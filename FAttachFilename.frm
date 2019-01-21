VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmAttachFilename 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Attach File"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   1005
   ClientWidth     =   8760
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFile 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   8295
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   150
         TabIndex        =   0
         Tag             =   "00-File Name (Do not Enter Extension TXT)"
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1620
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   9
      Top             =   4545
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
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
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "Attach File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Tag             =   "Select the Employee listed above"
         Top             =   150
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Tag             =   "Print the Employee Listing"
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   5040
      MultiSelect     =   2  'Extended
      Pattern         =   "*.doc;*.xls;*.ppt;*.pdf;*.jpg;*.docx"
      TabIndex        =   3
      Top             =   690
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   3465
      Left            =   2100
      TabIndex        =   2
      Top             =   1020
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   690
      Width           =   2895
   End
   Begin VB.Frame frAdditionalInfo 
      Caption         =   "K. Additional Information"
      Height          =   3255
      Left            =   120
      TabIndex        =   100
      Top             =   6000
      Width           =   8535
      Begin VB.TextBox txtAdditionalInfo 
         Appearance      =   0  'Flat
         DataField       =   "F7_ADDITIONAL_INFO"
         Height          =   2535
         Left            =   120
         MaxLength       =   4000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   101
         Tag             =   "00-Additional Information"
         Top             =   480
         Width           =   8205
      End
   End
   Begin VB.Frame frWorkSchedule 
      Caption         =   "I. Work Schedule (Complete either A, Bor C. Do not include overtime shifts)"
      Height          =   4695
      Left            =   120
      TabIndex        =   50
      Top             =   5640
      Width           =   8535
      Begin VB.OptionButton optWorkSchedule 
         Caption         =   "(C.) Varied or Irregular Work Schedule - Provide the total number of regular hours and shifts for each week for the"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   53
         Tag             =   "40-Varied or Irregular Work Schedule"
         Top             =   2280
         Width           =   8295
      End
      Begin VB.OptionButton optWorkSchedule 
         Caption         =   "(B.) Repeating Rotational Shift Worker - Provide"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Tag             =   "40-Repeating Rotational Shift Worker"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.OptionButton optWorkSchedule 
         Caption         =   "(A.) Regular Schedule - Indicate normal work days and hours."
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Tag             =   "40-Regular Work Schedule"
         Top             =   360
         Width           =   4695
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_MON"
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   55
         Tag             =   "11-Work Schedule for Monday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_TUE"
         Height          =   285
         Index           =   2
         Left            =   2640
         TabIndex        =   56
         Tag             =   "11-Work Schedule for Tuesday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_WED"
         Height          =   285
         Index           =   3
         Left            =   3600
         TabIndex        =   57
         Tag             =   "11-Work Schedule for Wednesday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_FRI"
         Height          =   285
         Index           =   5
         Left            =   5520
         TabIndex        =   58
         Tag             =   "11-Work Schedule for Friday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_THU"
         Height          =   285
         Index           =   4
         Left            =   4560
         TabIndex        =   59
         Tag             =   "11-Work Schedule for Thursday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_SAT"
         Height          =   285
         Index           =   6
         Left            =   6480
         TabIndex        =   60
         Tag             =   "11-Work Schedule for Saturday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_SUN"
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   67
         Tag             =   "11-Work Schedule for Sunday"
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medNoDayOn 
         DataField       =   "F7_NUM_DAYS_ON"
         Height          =   285
         Left            =   1695
         TabIndex        =   73
         Tag             =   "11-Number of Days On"
         Top             =   1800
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medNoDayOff 
         DataField       =   "F7_NUM_DAYS_OFF"
         Height          =   285
         Left            =   3495
         TabIndex        =   74
         Tag             =   "11-Number of Days Off"
         Top             =   1800
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medHrsShift 
         DataField       =   "F7_HRS_SHIFT"
         Height          =   285
         Left            =   5325
         TabIndex        =   75
         Tag             =   "11-Hours per Shift(s)"
         Top             =   1800
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medNoWksCycle 
         DataField       =   "F7_NUM_WKS_CYCLE"
         Height          =   285
         Left            =   7740
         TabIndex        =   76
         Tag             =   "11-Number of Weeks in Cycle"
         Top             =   1800
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin INFOHR_Controls.DateLookup dlpWk1FDate 
         DataField       =   "F7_FWEEK1"
         Height          =   285
         Left            =   2280
         TabIndex        =   84
         Tag             =   "41-Week 1 From Date"
         Top             =   3075
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk1TDate 
         DataField       =   "F7_TWEEK2"
         Height          =   285
         Left            =   3960
         TabIndex        =   85
         Tag             =   "41-Week 1 To Date"
         Top             =   3075
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk2FDate 
         DataField       =   "F7_FWEEK2"
         Height          =   285
         Left            =   5640
         TabIndex        =   86
         Tag             =   "41-Week 2 From Date"
         Top             =   3075
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk2TDate 
         DataField       =   "F7_TWEEK1"
         Height          =   285
         Left            =   7320
         TabIndex        =   87
         Tag             =   "41-Week 2 To Date"
         Top             =   3075
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk3FDate 
         DataField       =   "F7_FWEEK3"
         Height          =   285
         Left            =   2280
         TabIndex        =   88
         Tag             =   "41-Week 3 From Date"
         Top             =   3600
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk3TDate 
         DataField       =   "F7_TWEEK3"
         Height          =   285
         Left            =   3960
         TabIndex        =   89
         Tag             =   "41-Week 3 To Date"
         Top             =   3600
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk4FDate 
         DataField       =   "F7_FWEEK4"
         Height          =   285
         Left            =   5640
         TabIndex        =   90
         Tag             =   "41-Week 4 From Date"
         Top             =   3600
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk4TDate 
         DataField       =   "F7_TWEEK4"
         Height          =   285
         Left            =   7320
         TabIndex        =   91
         Tag             =   "41-Week 4 To Date"
         Top             =   3600
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin MSMask.MaskEdBox medTotShiftsWrkWK1 
         DataField       =   "F7_TOT_SHIFT_WEEK1"
         Height          =   285
         Left            =   2280
         TabIndex        =   92
         Tag             =   "11-Total Shifts Worked Week 1"
         Top             =   4320
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotShiftsWrkWK2 
         DataField       =   "F7_TOT_SHIFT_WEEK2"
         Height          =   285
         Left            =   3240
         TabIndex        =   93
         Tag             =   "11-Total Shifts Worked Week 2"
         Top             =   4320
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotShiftsWrkWK3 
         DataField       =   "F7_TOT_SHIFT_WEEK3"
         Height          =   285
         Left            =   4200
         TabIndex        =   94
         Tag             =   "11-Total Shifts Worked Week 3"
         Top             =   4320
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotShiftsWrkWK4 
         DataField       =   "F7_TOT_SHIFT_WEEK4"
         Height          =   285
         Left            =   5160
         TabIndex        =   95
         Tag             =   "11-Total Shifts Worked Week 4"
         Top             =   4320
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotHrsWrkWK1 
         DataField       =   "F7_TOT_HRS_WEEK1"
         Height          =   285
         Left            =   2280
         TabIndex        =   96
         Tag             =   "11-Total Hours Worked Week 1"
         Top             =   3960
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotHrsWrkWK2 
         DataField       =   "F7_TOT_HRS_WEEK2"
         Height          =   285
         Left            =   3240
         TabIndex        =   97
         Tag             =   "11-Total Hours Worked Week 2"
         Top             =   3960
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotHrsWrkWK3 
         DataField       =   "F7_TOT_HRS_WEEK3"
         Height          =   285
         Left            =   4200
         TabIndex        =   98
         Tag             =   "11-Total Hours Worked Week 3"
         Top             =   3960
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox medTotHrsWrkWK4 
         DataField       =   "F7_TOT_HRS_WEEK4"
         Height          =   285
         Left            =   5160
         TabIndex        =   99
         Tag             =   "11-Total Hours Worked Week 4"
         Top             =   3960
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 4"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   7080
         TabIndex        =   83
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   3720
         TabIndex        =   82
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   7080
         TabIndex        =   81
         Top             =   2780
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   3720
         TabIndex        =   80
         Top             =   2780
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Shifts Worked"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   720
         TabIndex        =   79
         Top             =   4365
         Width           =   1410
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Hours Worked"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   720
         TabIndex        =   78
         Top             =   4005
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "From / To Dates"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   720
         TabIndex        =   77
         Top             =   3120
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF WEEKS IN CYCLE"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   16
         Left            =   6120
         TabIndex        =   72
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "HOURS PER SHIFT(s)"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   15
         Left            =   4320
         TabIndex        =   71
         Top             =   1680
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF DAYS OFF"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   14
         Left            =   2520
         TabIndex        =   70
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF DAYS ON"
         ForeColor       =   &H80000008&
         Height          =   435
         Index           =   13
         Left            =   720
         TabIndex        =   69
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sunday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   720
         TabIndex        =   68
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   1680
         TabIndex        =   66
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tuesday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   2640
         TabIndex        =   65
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Wednesday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   3555
         TabIndex        =   64
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Friday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   63
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Thursday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   4560
         TabIndex        =   62
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Saturday"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   6480
         TabIndex        =   61
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4 weeks prior to the accident/illness. (Do not include overtime hours or shifts here)."
         Height          =   195
         Left            =   3240
         TabIndex        =   54
         Top             =   2535
         Width           =   5835
      End
   End
   Begin VB.Frame frAdditionalWage 
      Caption         =   "H. Additional Wage Information"
      Height          =   9135
      Left            =   120
      TabIndex        =   102
      Top             =   5280
      Width           =   8535
      Begin VB.OptionButton optBeingPaidYN 
         Caption         =   "No"
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   223
         Tag             =   "40-Is the worker being paid while he/dshe recovers? No"
         Top             =   2955
         Width           =   615
      End
      Begin VB.OptionButton optBeingPaidYN 
         Caption         =   "Yes"
         Height          =   285
         Index           =   0
         Left            =   5520
         TabIndex        =   222
         Tag             =   "40-Is the worker being paid while he/dshe recovers? Yes"
         Top             =   2955
         Width           =   615
      End
      Begin VB.TextBox txtOtherEarnings4 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_4"
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   220
         Top             =   8520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOtherEarnings3 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_3"
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   219
         Top             =   8160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOtherEarnings2 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_2"
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   218
         Top             =   7800
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOtherEarnings1 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_1"
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   217
         Top             =   7440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox comOtherEarnings4 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FAttachFilename.frx":0000
         Left            =   6600
         List            =   "FAttachFilename.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   216
         Tag             =   "10-Other Earnings"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.ComboBox comOtherEarnings3 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FAttachFilename.frx":0004
         Left            =   5040
         List            =   "FAttachFilename.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   215
         Tag             =   "10-Other Earnings"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.ComboBox comOtherEarnings2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FAttachFilename.frx":0008
         Left            =   3360
         List            =   "FAttachFilename.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   214
         Tag             =   "10-Other Earnings"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.ComboBox comOtherEarnings1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FAttachFilename.frx":000C
         Left            =   1800
         List            =   "FAttachFilename.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   213
         Tag             =   "10-Other Earnings"
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox txtBeingPaidOther 
         Appearance      =   0  'Flat
         DataField       =   "F7_WORKER_OTHER"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4200
         MaxLength       =   25
         TabIndex        =   138
         Tag             =   "01-Worker being paid while he/dshe recovers, Other"
         Top             =   3240
         Width           =   3555
      End
      Begin VB.OptionButton optFullRegOther 
         Caption         =   "Other"
         Height          =   285
         Index           =   1
         Left            =   3360
         TabIndex        =   137
         Tag             =   "40-Is the worker being paid while he/dshe recovers? Other"
         Top             =   3270
         Width           =   1215
      End
      Begin VB.OptionButton optFullRegOther 
         Caption         =   "Full/Regular"
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   136
         Tag             =   "40-Is the worker being paid while he/dshe recovers? Full/Regular"
         Top             =   3270
         Width           =   1215
      End
      Begin VB.TextBox txtHourLastWorked 
         Appearance      =   0  'Flat
         DataField       =   "F7_LAST_WORK_TIME"
         Height          =   285
         Left            =   5280
         MaxLength       =   5
         TabIndex        =   133
         Tag             =   "00-Normal From Working Time"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtNormTimeTo 
         Appearance      =   0  'Flat
         DataField       =   "F7_LAST_DAY_WORK_TTIME"
         Height          =   285
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   130
         Tag             =   "00-Normal To Working Time"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtNormTimeFrom 
         Appearance      =   0  'Flat
         DataField       =   "F7_LAST_DAY_WORK_FTIME"
         Height          =   285
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   129
         Tag             =   "00-Normal From Working Time"
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton optNormFTimeLastWorkYN 
         Caption         =   "PM"
         Height          =   285
         Index           =   2
         Left            =   5760
         TabIndex        =   122
         Tag             =   "40-Last Normal Worked To Time, PM"
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton optNormTTimeLastWorkYN 
         Caption         =   "AM"
         Height          =   285
         Index           =   0
         Left            =   5040
         TabIndex        =   121
         Tag             =   "40-Last Normal Worked To Time, AM"
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton optNormFTimeLastWorkYN 
         Caption         =   "PM"
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   120
         Tag             =   "40-Last Normal Worked From Time, PM"
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton optNormFTimeLastWorkYN 
         Caption         =   "AM"
         Height          =   285
         Index           =   0
         Left            =   5040
         TabIndex        =   119
         Tag             =   "40-Last Normal Worked From Time, AM"
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton optTimeLastWorkYN 
         Caption         =   "PM"
         Height          =   285
         Index           =   1
         Left            =   6960
         TabIndex        =   118
         Tag             =   "40-Last Worked Time, PM"
         Top             =   1080
         Width           =   615
      End
      Begin VB.OptionButton optTimeLastWorkYN 
         Caption         =   "AM"
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   117
         Tag             =   "40-Last Worked Time, AM"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtFedCodeAmt 
         Appearance      =   0  'Flat
         DataField       =   "F7_FED_AMT"
         Height          =   285
         Left            =   3480
         TabIndex        =   112
         Tag             =   "00-Federal Code/Amount"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtProvCodeAmt 
         Appearance      =   0  'Flat
         DataField       =   "F7_PROV_AMT"
         Height          =   285
         Left            =   5640
         TabIndex        =   111
         Tag             =   "00- Provincial Code/Amount"
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optVacPerctYN 
         Caption         =   "No"
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   109
         Tag             =   "40-Vacation pay on each cheque? No"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optVacPerctYN 
         Caption         =   "Yes"
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   108
         Tag             =   "40-Vacation pay on each cheque? Yes"
         Top             =   720
         Width           =   615
      End
      Begin MSMask.MaskEdBox medVacPerct 
         DataField       =   "F7_VACPC"
         Height          =   285
         Left            =   6360
         TabIndex        =   110
         Tag             =   "11-Vacation Pay Percentage"
         Top             =   720
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin INFOHR_Controls.DateLookup dlpDateLastWork 
         DataField       =   "F7_LAST_WORK_DATE"
         Height          =   285
         Left            =   3120
         TabIndex        =   123
         Tag             =   "41-Date Last Worked"
         Top             =   1080
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin MSMask.MaskEdBox medLastActualEarnings 
         DataField       =   "F7_LAST_DAY_ACT_EARN"
         Height          =   285
         Left            =   3240
         TabIndex        =   125
         Tag             =   "20-Actual earnings for last day worked"
         Top             =   2160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLastNormalEarnings 
         DataField       =   "F7_LAST_DAY_NORM_EARN"
         Height          =   285
         Left            =   3240
         TabIndex        =   127
         Tag             =   "20-Normal earnings for last day worked"
         Top             =   2520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK1 
         Height          =   285
         Left            =   1560
         TabIndex        =   150
         Tag             =   "41-Other Earning From Date Week 1"
         Top             =   5160
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK2 
         Height          =   285
         Left            =   1560
         TabIndex        =   151
         Tag             =   "41-Other Earning From Date Week 2"
         Top             =   5520
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK3 
         Height          =   285
         Left            =   1560
         TabIndex        =   152
         Tag             =   "41-Other Earning From Date Week 3"
         Top             =   5880
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK4 
         Height          =   285
         Left            =   1560
         TabIndex        =   153
         Tag             =   "41-Other Earning From Date Week 4"
         Top             =   6240
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK1 
         Height          =   285
         Left            =   3240
         TabIndex        =   154
         Tag             =   "41-Other Earning To Date Week 1"
         Top             =   5160
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK2 
         Height          =   285
         Left            =   3240
         TabIndex        =   155
         Tag             =   "41-Other Earning To Date Week 2"
         Top             =   5520
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK3 
         Height          =   285
         Left            =   3240
         TabIndex        =   156
         Tag             =   "41-Other Earning To Date Week 3"
         Top             =   5880
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK4 
         Height          =   285
         Left            =   3240
         TabIndex        =   157
         Tag             =   "41-Other Earning To Date Week 4"
         Top             =   6240
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK1 
         DataField       =   "F7_MAND_OVT_PAY_WK1"
         Height          =   285
         Left            =   5040
         TabIndex        =   158
         Tag             =   "20-Mandatory Overtime Pay Week 1"
         Top             =   5160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK2 
         DataField       =   "F7_MAND_OVT_PAY_WK2"
         Height          =   285
         Left            =   5040
         TabIndex        =   160
         Tag             =   "20-Mandatory Overtime Pay Week 2"
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK3 
         DataField       =   "F7_MAND_OVT_PAY_WK3"
         Height          =   285
         Left            =   5040
         TabIndex        =   162
         Tag             =   "20-Mandatory Overtime Pay Week 3"
         Top             =   5880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK4 
         DataField       =   "F7_MAND_OVT_PAY_WK4"
         Height          =   285
         Left            =   5040
         TabIndex        =   164
         Tag             =   "20-Mandatory Overtime Pay Week 4"
         Top             =   6240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK1 
         DataField       =   "F7_VOL_OVT_PAY_WK1"
         Height          =   285
         Left            =   6600
         TabIndex        =   166
         Tag             =   "20-Voluntary Overtime Pay Week 1"
         Top             =   5160
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK2 
         DataField       =   "F7_VOL_OVT_PAY_WK2"
         Height          =   285
         Left            =   6600
         TabIndex        =   168
         Tag             =   "20-Voluntary Overtime Pay Week 2"
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK3 
         DataField       =   "F7_VOL_OVT_PAY_WK3"
         Height          =   285
         Left            =   6600
         TabIndex        =   170
         Tag             =   "20-Voluntary Overtime Pay Week 3"
         Top             =   5880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK4 
         DataField       =   "F7_VOL_OVT_PAY_WK4"
         Height          =   285
         Left            =   6600
         TabIndex        =   172
         Tag             =   "20-Voluntary Overtime Pay Week 4"
         Top             =   6240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK1 
         DataField       =   "F7_OTH_EARN_1_WK1"
         Height          =   285
         Left            =   1800
         TabIndex        =   176
         Tag             =   "20-Other Earnings 1 - Week 1"
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK2 
         DataField       =   "F7_OTH_EARN_1_WK2"
         Height          =   285
         Left            =   1800
         TabIndex        =   178
         Tag             =   "20-Other Earnings 1 - Week 2"
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK3 
         DataField       =   "F7_OTH_EARN_1_WK3"
         Height          =   285
         Left            =   1800
         TabIndex        =   180
         Tag             =   "20-Other Earnings 1 - Week 3"
         Top             =   8280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK4 
         DataField       =   "F7_OTH_EARN_1_WK4"
         Height          =   285
         Left            =   1800
         TabIndex        =   182
         Tag             =   "20-Other Earnings 1 - Week 4"
         Top             =   8640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK1 
         DataField       =   "F7_OTH_EARN_2_WK1"
         Height          =   285
         Left            =   3360
         TabIndex        =   184
         Tag             =   "20-Other Earnings 2 - Week 1"
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK2 
         DataField       =   "F7_OTH_EARN_2_WK2"
         Height          =   285
         Left            =   3360
         TabIndex        =   186
         Tag             =   "20-Other Earnings 2 - Week 2"
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK3 
         DataField       =   "F7_OTH_EARN_2_WK3"
         Height          =   285
         Left            =   3360
         TabIndex        =   188
         Tag             =   "20-Other Earnings 2 - Week 3"
         Top             =   8280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK4 
         DataField       =   "F7_OTH_EARN_2_WK4"
         Height          =   285
         Left            =   3360
         TabIndex        =   190
         Tag             =   "20-Other Earnings 2 - Week 4"
         Top             =   8640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK1 
         DataField       =   "F7_OTH_EARN_3_WK1"
         Height          =   285
         Left            =   5040
         TabIndex        =   192
         Tag             =   "20-Other Earnings 3 - Week 1"
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK2 
         DataField       =   "F7_OTH_EARN_3_WK2"
         Height          =   285
         Left            =   5040
         TabIndex        =   194
         Tag             =   "20-Other Earnings 3 - Week 2"
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK3 
         DataField       =   "F7_OTH_EARN_3_WK3"
         Height          =   285
         Left            =   5040
         TabIndex        =   196
         Tag             =   "20-Other Earnings 3 - Week 3"
         Top             =   8280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK4 
         DataField       =   "F7_OTH_EARN_3_WK4"
         Height          =   285
         Left            =   5040
         TabIndex        =   198
         Tag             =   "20-Other Earnings 3 - Week 4"
         Top             =   8640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK1 
         DataField       =   "F7_OTH_EARN_4_WK1"
         Height          =   285
         Left            =   6600
         TabIndex        =   200
         Tag             =   "20-Other Earnings 4 - Week 1"
         Top             =   7560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK2 
         DataField       =   "F7_OTH_EARN_4_WK2"
         Height          =   285
         Left            =   6600
         TabIndex        =   202
         Tag             =   "20-Other Earnings 4 - Week 2"
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK3 
         DataField       =   "F7_OTH_EARN_4_WK3"
         Height          =   285
         Left            =   6600
         TabIndex        =   204
         Tag             =   "20-Other Earnings 4 - Week 3"
         Top             =   8280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK4 
         DataField       =   "F7_OTH_EARN_4_WK4"
         Height          =   285
         Left            =   6600
         TabIndex        =   206
         Tag             =   "20-Other Earnings 4 - Week 4"
         Top             =   8640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   195
         Left            =   4800
         TabIndex        =   225
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   195
         Left            =   2640
         TabIndex        =   224
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Is the worker being paid while he/she recovers?"
         Height          =   195
         Left            =   1920
         TabIndex        =   221
         Top             =   3000
         Width           =   3405
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 4"
         Height          =   195
         Left            =   720
         TabIndex        =   212
         Top             =   8685
         Width           =   570
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 3"
         Height          =   195
         Left            =   720
         TabIndex        =   211
         Top             =   8325
         Width           =   570
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 2"
         Height          =   195
         Left            =   720
         TabIndex        =   210
         Top             =   7965
         Width           =   570
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 1"
         Height          =   195
         Left            =   720
         TabIndex        =   209
         Top             =   7605
         Width           =   570
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         Height          =   195
         Left            =   720
         TabIndex        =   208
         Top             =   7200
         Width           =   450
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   207
         Top             =   8685
         Width           =   90
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   205
         Top             =   8325
         Width           =   90
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   203
         Top             =   7965
         Width           =   90
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   201
         Top             =   7605
         Width           =   90
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   199
         Top             =   8685
         Width           =   90
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   197
         Top             =   8325
         Width           =   90
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   195
         Top             =   7965
         Width           =   90
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   193
         Top             =   7605
         Width           =   90
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   3240
         TabIndex        =   191
         Top             =   8685
         Width           =   90
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   3240
         TabIndex        =   189
         Top             =   8325
         Width           =   90
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   3240
         TabIndex        =   187
         Top             =   7965
         Width           =   90
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   3240
         TabIndex        =   185
         Top             =   7605
         Width           =   90
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   1680
         TabIndex        =   183
         Top             =   8685
         Width           =   90
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   1680
         TabIndex        =   181
         Top             =   8325
         Width           =   90
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   1680
         TabIndex        =   179
         Top             =   7965
         Width           =   90
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   1680
         TabIndex        =   177
         Top             =   7605
         Width           =   90
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Use these spaces for any other earnings (indicate Commission, Differentials, Premiums, Bonus, Tips, In Lieu %, etc.)."
         Height          =   435
         Left            =   720
         TabIndex        =   175
         Top             =   6720
         Width           =   7185
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   $"FAttachFilename.frx":0010
         Height          =   435
         Left            =   720
         TabIndex        =   174
         Top             =   4320
         Width           =   7200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   173
         Top             =   6285
         Width           =   90
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   171
         Top             =   5925
         Width           =   90
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   169
         Top             =   5565
         Width           =   90
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   6480
         TabIndex        =   167
         Top             =   5205
         Width           =   90
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   165
         Top             =   6285
         Width           =   90
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   163
         Top             =   5925
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   161
         Top             =   5565
         Width           =   90
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   4920
         TabIndex        =   159
         Top             =   5205
         Width           =   90
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 4"
         Height          =   195
         Left            =   720
         TabIndex        =   149
         Top             =   6285
         Width           =   570
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 3"
         Height          =   195
         Left            =   720
         TabIndex        =   148
         Top             =   5925
         Width           =   570
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 2"
         Height          =   195
         Left            =   720
         TabIndex        =   147
         Top             =   5565
         Width           =   570
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 1"
         Height          =   195
         Left            =   720
         TabIndex        =   146
         Top             =   5205
         Width           =   570
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voluntary"
         Height          =   195
         Left            =   6840
         TabIndex        =   145
         Top             =   4830
         Width           =   660
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mandatory"
         Height          =   195
         Left            =   5280
         TabIndex        =   144
         Top             =   4830
         Width           =   750
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         Height          =   195
         Left            =   3840
         TabIndex        =   143
         Top             =   4830
         Width           =   585
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         Height          =   195
         Left            =   2160
         TabIndex        =   142
         Top             =   4830
         Width           =   735
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
         Height          =   195
         Left            =   720
         TabIndex        =   141
         Top             =   4830
         Width           =   450
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provide the total of additional earnings for each week for the 4 weeks before the accident/illness."
         Height          =   195
         Left            =   480
         TabIndex        =   140
         Top             =   3960
         Width           =   6870
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8. Other Earnings (Not Regular Wages): "
         Height          =   195
         Left            =   120
         TabIndex        =   139
         Top             =   3720
         Width           =   2865
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If yes, indicate: "
         Height          =   195
         Left            =   480
         TabIndex        =   135
         Top             =   3285
         Width           =   1110
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7. Advances on wages:"
         Height          =   195
         Left            =   120
         TabIndex        =   134
         Top             =   3000
         Width           =   1680
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   3600
         TabIndex        =   132
         Top             =   1845
         Width           =   195
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   3600
         TabIndex        =   131
         Top             =   1485
         Width           =   345
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   3120
         TabIndex        =   128
         Top             =   2565
         Width           =   90
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         Height          =   195
         Left            =   3120
         TabIndex        =   126
         Top             =   2205
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Left            =   6960
         TabIndex        =   124
         Top             =   765
         Width           =   120
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6. Normal earnings for last day worked"
         Height          =   195
         Left            =   120
         TabIndex        =   116
         Top             =   2565
         Width           =   2700
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. Actual earnings for last day worked"
         Height          =   195
         Left            =   120
         TabIndex        =   115
         Top             =   2205
         Width           =   2655
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Normal working hours on last day worked"
         Height          =   195
         Left            =   120
         TabIndex        =   114
         Top             =   1485
         Width           =   3090
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Date and hour last worked:"
         Height          =   195
         Left            =   120
         TabIndex        =   113
         Top             =   1125
         Width           =   2100
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provide percentage"
         Height          =   195
         Left            =   4800
         TabIndex        =   107
         Top             =   765
         Width           =   1395
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Vacation pay on each cheque?"
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   765
         Width           =   2415
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincial"
         Height          =   195
         Left            =   4800
         TabIndex        =   105
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Federal"
         Height          =   195
         Left            =   2640
         TabIndex        =   104
         Top             =   405
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Net Claim Code of Amount:"
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   405
         Width           =   2085
      End
   End
   Begin VB.Frame frFReturnToWork 
      Caption         =   "F. Return To Work"
      Height          =   4335
      Left            =   120
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   8535
      Begin VB.CommandButton cmdImport1 
         Caption         =   "Import"
         Height          =   330
         Left            =   7560
         TabIndex        =   229
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton optResponsible 
         Caption         =   "Myself"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   47
         Tag             =   "40-Responsible for arranging worker's return to work Myself"
         Top             =   3600
         Width           =   855
      End
      Begin VB.OptionButton optResponsible 
         Caption         =   "Other"
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   46
         Tag             =   "40-Responsible for arranging worker's return to work Other"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox txtResponsibleName 
         Appearance      =   0  'Flat
         DataField       =   "F7_RESPONS_NAME"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   35
         TabIndex        =   45
         Tag             =   "01-Responsible for arranging worker's return to work"
         Top             =   3600
         Width           =   3555
      End
      Begin VB.TextBox txtResponsiblePhoneExt 
         Appearance      =   0  'Flat
         DataField       =   "F7_RESPONS_PHONE_EXT"
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   44
         Tag             =   "01-Phone Extension"
         Top             =   3960
         Width           =   795
      End
      Begin VB.CheckBox chkWrittenOfferAttached 
         Caption         =   "If Declined please attach a copy of the written offer given to the worker."
         DataField       =   "F7_DECLINE_ATTACHED"
         Height          =   195
         Left            =   1560
         TabIndex        =   43
         Tag             =   "Written Offer attachment"
         Top             =   2760
         Width           =   5685
      End
      Begin VB.OptionButton optAccDecl 
         Caption         =   "Accepted"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   42
         Tag             =   "40-Return to Work Offer, Accepted"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton optAccDecl 
         Caption         =   "Declined"
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   41
         Tag             =   "40-Return to Work Offer, Declined"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton optOffered 
         Caption         =   "Yes"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   39
         Tag             =   "40-Modified work been offered, Yes"
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton optOffered 
         Caption         =   "No"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   38
         Tag             =   "40-Modified work been offered, No"
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton optDiscussed 
         Caption         =   "Yes"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   36
         Tag             =   "40-Modified work been discussed, Yes"
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton optDiscussed 
         Caption         =   "No"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   35
         Tag             =   "40-Modified work been discussed, No"
         Top             =   1200
         Width           =   615
      End
      Begin VB.OptionButton optLimitations 
         Caption         =   "Yes"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Tag             =   "40-Provided with work limitations, Yes"
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optLimitations 
         Caption         =   "No"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   32
         Tag             =   "40-Provided with work limitations, No"
         Top             =   600
         Width           =   615
      End
      Begin MSMask.MaskEdBox medResponsiblePhone 
         DataField       =   "F7_RESPONS_PHONE"
         Height          =   285
         Left            =   3120
         TabIndex        =   226
         Tag             =   "10-Telephone Number"
         Top             =   3960
         Width           =   1485
         _ExtentX        =   2619
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
      Begin VB.Image imgNoSec1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   7200
         Picture         =   "FAttachFilename.frx":00C8
         Top             =   3045
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport1 
         Alignment       =   1  'Right Justify
         Caption         =   "Written Offer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   5280
         TabIndex        =   230
         Top             =   3075
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ext."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   4695
         TabIndex        =   228
         Top             =   4005
         Width           =   270
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   2265
         TabIndex        =   227
         Top             =   4005
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Who is responsible for arranging worker's return to work"
         Height          =   195
         Left            =   120
         TabIndex        =   49
         Top             =   3360
         Width           =   4095
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   48
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If yes, wast it "
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Has modified work been offered to this worker?"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   3510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Has modified work been discussed with this worker?"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   3870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Have you been provided with work limitations for this worker's injury?"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   5010
      End
      Begin VB.Image imgSec1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   7200
         Picture         =   "FAttachFilename.frx":0212
         Top             =   3045
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Frame frELostTime 
      Caption         =   "E. Lost Time - No Lost Time"
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   8535
      Begin VB.TextBox txtConfirmPhoneExt 
         Appearance      =   0  'Flat
         DataField       =   "F7_CONFIRM_PHONE_EXT"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "01-Phone Extension"
         Top             =   3240
         Width           =   795
      End
      Begin VB.TextBox txtConfirmedByName 
         Appearance      =   0  'Flat
         DataField       =   "F7_CONFIRM_NAME"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         MaxLength       =   35
         TabIndex        =   24
         Tag             =   "01-Information was confirmed by"
         Top             =   2880
         Width           =   3555
      End
      Begin VB.OptionButton optConfirmedBy 
         Caption         =   "Other"
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   23
         Tag             =   "40-Employee's Premises Yes"
         Top             =   2880
         Width           =   855
      End
      Begin VB.OptionButton optConfirmedBy 
         Caption         =   "Myself"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Tag             =   "40-Employee's Premises Yes"
         Top             =   2880
         Width           =   855
      End
      Begin VB.OptionButton optRegMod 
         Caption         =   "modified work"
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   18
         Tag             =   "40-Modified Work"
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton optRegMod 
         Caption         =   "regular work"
         Height          =   285
         Index           =   0
         Left            =   5400
         TabIndex        =   17
         Tag             =   "40-Regular Work"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optLostTime 
         Caption         =   "Has lost time and/or earnings. (Complete ALL remaining sections)."
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Tag             =   "40-Lost time and/or earnings"
         Top             =   1440
         Width           =   5535
      End
      Begin VB.OptionButton optLostTime 
         Caption         =   "Returned to his/her regular job and has not lost any time and/or earnings. (Complete sections G and J)."
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Tag             =   "40-Returned to regular job and has not lost any time and/or earnings"
         Top             =   720
         Width           =   7695
      End
      Begin VB.OptionButton optLostTime 
         Caption         =   "Returned to modified work and has not lost any time and/or earnings. (Complete sections F, G, and J)."
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Tag             =   "40-Returned to modified work and has not lost any time and/or earnings"
         Top             =   1080
         Width           =   7575
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "F7_LOST_DATE"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   15
         Tag             =   "41-Date worker first lost time"
         Top             =   1800
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "F7_RETURN_DATE"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   16
         Tag             =   "41-Date worker returned to work (if known)"
         Top             =   2160
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin MSMask.MaskEdBox medConfirmPhone 
         DataField       =   "F7_CONFIRM_PHONE"
         Height          =   285
         Left            =   2970
         TabIndex        =   25
         Tag             =   "10-Alternate Telephone Number"
         Top             =   3240
         Width           =   1485
         _ExtentX        =   2619
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
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2430
         TabIndex        =   29
         Top             =   2925
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ext."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   4575
         TabIndex        =   28
         Top             =   3285
         Width           =   270
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   2085
         TabIndex        =   26
         Top             =   3285
         Width           =   765
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "2. This Lost Time - No Lost Time - Modified Work information was confirmed by:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   2640
         Width           =   5700
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date worker returned to work (if known)"
         Height          =   195
         Left            =   840
         TabIndex        =   20
         Top             =   2205
         Width           =   2940
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Provide date worker first lost time"
         Height          =   195
         Left            =   840
         TabIndex        =   19
         Top             =   1845
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Please choose one of the following indicators. After the day of accident/awareness of illness, this worker:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   7590
      End
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "File to Attach"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   660
      Width           =   1620
   End
End
Attribute VB_Name = "frmAttachFilename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FPath ', UPDTCNT
Dim AttachFile As String

Private Sub cmdClose_Click()
    glbDocName = ""
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Sub cmdModify_Click()
    On Error GoTo Mod_Err
    
    Screen.MousePointer = HOURGLASS
    
    If Not chkDoc() Then Exit Sub
    
    frmJobDocument.txtFileName.Text = AttachFile
    
    Screen.MousePointer = DEFAULT
    
    Close
    
    Unload Me
    
Exit Sub


Mod_Err:
    If Err = 53 Then Resume Next
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Attach File", "Update")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub Dir1_Change()
    ChDir Dir1.Path
    File1.Path = Dir1.Path
    'File1.Pattern = "*.JPG"
End Sub

Private Sub Drive1_Change()
    Dim xdir, xerror
    On Error GoTo CKERROR
    
    xerror = False
    Dir1.Path = Drive1.Drive
    
Exit Sub
CKERROR:
    If Err = 68 Then
         MsgBox "Invalid Drive Selected"
         Drive1.Drive = App.Path
         xerror = True
         Resume Next
    End If
    MsgBox "ERROR " & Str(Err)
    xerror = True
    Resume Next
End Sub

Private Sub File1_Click()
    Dim iit As Integer
    Dim ii1 As Long
    Dim sit As String
    For iit = 0 To File1.ListCount - 1
        If File1.selected(iit) Then
            sit = File1.List(iit)
            txtFileName.Text = UCase(File1.List(iit))
        End If
    Next
End Sub

Private Sub Form_Activate()
    Call INI_Controls(Me)
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    Dim rsEMP As New ADODB.Recordset
    Dim x%, SQLQ
    Dim Y%
    
    glbOnTop = "FRMATTACHFILENAME"
    Screen.MousePointer = HOURGLASS
    
    Drive1.Drive = "c:"
    Dir1.Path = "c:\"
    FPath = Dir1.Path
    
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


Sub txtFileName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub txtFileName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Function chkDoc()
    Dim Alphabet, xlen, I%, xwk, xok
    chkDoc = False
    On Error GoTo chkDoc_Err

    Screen.MousePointer = DEFAULT

    If Len(txtFileName) = 0 Then
        MsgBox "File Name is required."
        txtFileName.SetFocus
        Exit Function
    End If
    
    txtFileName = LTrim(txtFileName)
    xlen = Len(txtFileName)
    ' dkostka - 10/16/2001 - Added space and -_()! to end of alphabet, filenames can have these chars
    'Hemu - Ticket #16031 - With French accents - àâäæçéèêëîïôœùûü«€ÀÂÄÆÇÉÈÊËÎÏÔŒÙÛÜ»
    Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_()!., àâäæçéèêëîïôœùûü«€ÀÂÄÆÇÉÈÊËÎÏÔŒÙÛÜ»"

    xok = True
    For I% = 1 To xlen
        xwk = Mid(txtFileName, I%, 1)
        If InStr(Alphabet, xwk) = 0 Then
            xok = False
            Exit For
        End If
    Next
    If Not xok Then
        MsgBox "Invalid File Name"
        txtFileName.SetFocus
        Exit Function
    End If

    AttachFile = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\")) & txtFileName
    'MsAttachFilegBox AttachFile
    If Dir(AttachFile) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & AttachFile & "]"
        txtFileName.SetFocus
        Exit Function
    End If

    txtFileName.Text = AttachFile
    
chkDoc = True

Exit Function
chkDoc_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDoc", "Attach File", "edit/Add")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Function

Public Property Get ChangeAction() As UpdateStateEnum
    ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
    RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
    UpdateRight = True
End Property

Public Property Get Addable() As Boolean
    Addable = False
End Property

Public Property Get Updateble() As Boolean
    Updateble = True
End Property

Public Property Get Deleteble() As Boolean
    Deleteble = False
End Property

Public Property Get Printable() As Boolean
    Printable = False
End Property

Public Sub SET_UP_MODE()
    Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

