VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEInjF7Sections 
   Appearance      =   0  'Flat
   Caption         =   "WSIB Form 7 Sections: E, F, H, I and K"
   ClientHeight    =   10560
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   10410
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10560
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Tag             =   "01-Employee ID in the Division"
   Begin VB.Frame frJFilledBy 
      BorderStyle     =   0  'None
      Caption         =   "J. Form 7 Filled By"
      Height          =   2535
      Left            =   360
      TabIndex        =   268
      Top             =   1920
      Width           =   8775
      Begin VB.ComboBox cmbFilledByName 
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
         Left            =   1920
         TabIndex        =   271
         Tag             =   "11-Choose Name of Person Filling Form 7"
         Top             =   420
         Width           =   5175
      End
      Begin VB.TextBox txtFilledByName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "F7_NAME"
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
         Left            =   7440
         TabIndex        =   270
         Top             =   435
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Frame frFilledByDetails 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   272
         Top             =   720
         Width           =   7215
         Begin VB.TextBox txtFilledByTitle 
            Appearance      =   0  'Flat
            DataField       =   "F7_TITLE"
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   273
            Tag             =   "00-Official Title of the Person Completing the Form 7"
            Top             =   120
            Width           =   5175
         End
         Begin MSMask.MaskEdBox medFilledByTelephone 
            DataField       =   "F7_PHONE"
            Height          =   285
            Left            =   1800
            TabIndex        =   274
            Tag             =   "11-Telephone Number of the Person Completing Form 7"
            Top             =   540
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   27
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "(###) ###-####  Ext(######)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblPhone 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            TabIndex        =   276
            Top             =   585
            Width           =   465
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Official Title"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   275
            Top             =   165
            Width           =   1050
         End
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of person completing this report"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   269
         Top             =   480
         Width           =   795
      End
   End
   Begin VB.Frame frKAdditionalInfo 
      BorderStyle     =   0  'None
      Caption         =   "K. Additional Information"
      Height          =   6255
      Left            =   360
      TabIndex        =   229
      Top             =   2520
      Width           =   9615
      Begin VB.TextBox txtAdditionalInfo 
         Appearance      =   0  'Flat
         DataField       =   "F7_ADDITIONAL_INFO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   120
         MaxLength       =   3564
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   116
         Tag             =   "00-Additional Information"
         Top             =   480
         Width           =   9285
      End
   End
   Begin VB.Frame frHAdditionalWage 
      BorderStyle     =   0  'None
      Caption         =   "H. Additional Wage Information"
      Height          =   8175
      Left            =   360
      TabIndex        =   141
      Top             =   1635
      Width           =   9615
      Begin VB.Frame Frame14 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1800
         TabIndex        =   267
         Top             =   2550
         Width           =   2295
         Begin VB.OptionButton optFullRegOther 
            Caption         =   "Other"
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
            Left            =   1440
            TabIndex        =   48
            Tag             =   "40-Is the worker being paid while he/dshe recovers? Other"
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optFullRegOther 
            Caption         =   "Full/Regular"
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
            Left            =   0
            TabIndex        =   47
            Tag             =   "40-Is the worker being paid while he/dshe recovers? Full/Regular"
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame Frame13 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   7800
         TabIndex        =   266
         Top             =   1215
         Width           =   1335
         Begin VB.OptionButton optNormTTimeLastWorkAP 
            Caption         =   "PM"
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
            Left            =   720
            TabIndex        =   42
            Tag             =   "40-Last Normal Worked To Time, PM"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optNormTTimeLastWorkAP 
            Caption         =   "AM"
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
            Left            =   0
            TabIndex        =   41
            Tag             =   "40-Last Normal Worked To Time, AM"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5520
         TabIndex        =   265
         Top             =   2250
         Width           =   1335
         Begin VB.OptionButton optBeingPaidYN 
            Caption         =   "No"
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
            Left            =   720
            TabIndex        =   46
            Tag             =   "40-Is the worker being paid while he/dshe recovers? No"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optBeingPaidYN 
            Caption         =   "Yes"
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
            Left            =   0
            TabIndex        =   45
            Tag             =   "40-Is the worker being paid while he/dshe recovers? Yes"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   4920
         TabIndex        =   264
         Top             =   1215
         Width           =   1335
         Begin VB.OptionButton optNormFTimeLastWorkAP 
            Caption         =   "PM"
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
            Left            =   720
            TabIndex        =   39
            Tag             =   "40-Last Normal Worked From Time, PM"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optNormFTimeLastWorkAP 
            Caption         =   "AM"
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
            Left            =   0
            TabIndex        =   38
            Tag             =   "40-Last Normal Worked From Time, AM"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   6720
         TabIndex        =   263
         Top             =   855
         Width           =   1335
         Begin VB.OptionButton optTimeLastWorkAP 
            Caption         =   "PM"
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
            Left            =   720
            TabIndex        =   36
            Tag             =   "40-Last Worked Time, PM"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optTimeLastWorkAP 
            Caption         =   "AM"
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
            Left            =   0
            TabIndex        =   35
            Tag             =   "40-Last Worked Time, AM"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2760
         TabIndex        =   262
         Top             =   495
         Width           =   1335
         Begin VB.OptionButton optVacPerctYN 
            Caption         =   "No"
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
            Left            =   720
            TabIndex        =   31
            Tag             =   "40-Vacation pay on each cheque? No"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optVacPerctYN 
            Caption         =   "Yes"
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
            Left            =   0
            TabIndex        =   30
            Tag             =   "40-Vacation pay on each cheque? Yes"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtFullRegOther 
         Appearance      =   0  'Flat
         DataField       =   "F7_WORKER_FTREGOTHR"
         Height          =   285
         Left            =   7920
         MaxLength       =   5
         TabIndex        =   250
         Top             =   2520
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBeingPaidYN 
         Appearance      =   0  'Flat
         DataField       =   "FY_WORKER_PAID"
         Height          =   285
         Left            =   6960
         MaxLength       =   5
         TabIndex        =   249
         Top             =   2160
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNormTTimeLastWorkAP 
         Appearance      =   0  'Flat
         DataField       =   "F7_LAST_DAY_WORK_TAMPM"
         Height          =   285
         Left            =   8160
         MaxLength       =   5
         TabIndex        =   248
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtNormFTimeLastWorkAP 
         Appearance      =   0  'Flat
         DataField       =   "F7_LAST_DAY_WORK_FAMPM"
         Height          =   285
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   247
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTimeLastWorkAP 
         Appearance      =   0  'Flat
         DataField       =   "F7_LAST_WORK_AMPM"
         Height          =   285
         Left            =   8160
         MaxLength       =   5
         TabIndex        =   246
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtVacPerctYN 
         Appearance      =   0  'Flat
         DataField       =   "F7_VAC_PAY"
         Height          =   285
         Left            =   6840
         MaxLength       =   5
         TabIndex        =   245
         Top             =   480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBeingPaidOther 
         Appearance      =   0  'Flat
         DataField       =   "F7_WORKER_OTHER"
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
         Left            =   4200
         MaxLength       =   25
         TabIndex        =   49
         Tag             =   "01-Worker being paid while he/dshe recovers, Other"
         Top             =   2550
         Width           =   3555
      End
      Begin VB.TextBox txtProvCodeAmt 
         Appearance      =   0  'Flat
         DataField       =   "F7_PROV_AMT"
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
         Left            =   5760
         TabIndex        =   29
         Tag             =   "00- Provincial Code/Amount"
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox txtFedCodeAmt 
         Appearance      =   0  'Flat
         DataField       =   "F7_FED_AMT"
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
         Left            =   3480
         TabIndex        =   28
         Tag             =   "00-Federal Code/Amount"
         Top             =   120
         Width           =   975
      End
      Begin VB.ComboBox comOtherEarnings1 
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
         ItemData        =   "frmEInjF7Sections.frx":0000
         Left            =   1800
         List            =   "frmEInjF7Sections.frx":0002
         TabIndex        =   66
         Tag             =   "10-Other Earnings"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.ComboBox comOtherEarnings2 
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
         ItemData        =   "frmEInjF7Sections.frx":0004
         Left            =   3450
         List            =   "frmEInjF7Sections.frx":0006
         TabIndex        =   67
         Tag             =   "10-Other Earnings"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.ComboBox comOtherEarnings3 
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
         ItemData        =   "frmEInjF7Sections.frx":0008
         Left            =   5040
         List            =   "frmEInjF7Sections.frx":000A
         TabIndex        =   68
         Tag             =   "10-Other Earnings"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.ComboBox comOtherEarnings4 
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
         ItemData        =   "frmEInjF7Sections.frx":000C
         Left            =   6600
         List            =   "frmEInjF7Sections.frx":000E
         TabIndex        =   69
         Tag             =   "10-Other Earnings"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.TextBox txtOtherEarnings1 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_1"
         Height          =   285
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   145
         Top             =   6480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOtherEarnings2 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_2"
         Height          =   285
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   144
         Top             =   6840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOtherEarnings3 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_3"
         Height          =   285
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   143
         Top             =   7200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOtherEarnings4 
         Appearance      =   0  'Flat
         DataField       =   "F7_OTH_EARN_4"
         Height          =   285
         Left            =   8640
         MaxLength       =   15
         TabIndex        =   142
         Top             =   7560
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox medVacPerct 
         DataField       =   "F7_VACPC"
         Height          =   285
         Left            =   5760
         TabIndex        =   32
         Tag             =   "11-Vacation Pay Percentage"
         Top             =   480
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
         Left            =   3240
         TabIndex        =   33
         Tag             =   "41-Date Last Worked"
         Top             =   840
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin MSMask.MaskEdBox medLastActualEarnings 
         DataField       =   "F7_LAST_DAY_ACT_EARN"
         Height          =   285
         Left            =   3480
         TabIndex        =   43
         Tag             =   "20-Actual earnings for last day worked"
         Top             =   1560
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
         Left            =   3480
         TabIndex        =   44
         Tag             =   "20-Normal earnings for last day worked"
         Top             =   1920
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
         DataField       =   "F7_OTH_EARN_FROM_WK1"
         Height          =   285
         Left            =   1560
         TabIndex        =   50
         Tag             =   "41-Other Earning From Date Week 1"
         Top             =   4320
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK2 
         DataField       =   "F7_OTH_EARN_FROM_WK2"
         Height          =   285
         Left            =   1560
         TabIndex        =   54
         Tag             =   "41-Other Earning From Date Week 2"
         Top             =   4680
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK3 
         DataField       =   "F7_OTH_EARN_FROM_WK3"
         Height          =   285
         Left            =   1560
         TabIndex        =   58
         Tag             =   "41-Other Earning From Date Week 3"
         Top             =   5040
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnFromWK4 
         DataField       =   "F7_OTH_EARN_FROM_WK4"
         Height          =   285
         Left            =   1560
         TabIndex        =   62
         Tag             =   "41-Other Earning From Date Week 4"
         Top             =   5400
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK1 
         DataField       =   "F7_OTH_EARN_TO_WK1"
         Height          =   285
         Left            =   3240
         TabIndex        =   51
         Tag             =   "41-Other Earning To Date Week 1"
         Top             =   4320
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK2 
         DataField       =   "F7_OTH_EARN_TO_WK2"
         Height          =   285
         Left            =   3240
         TabIndex        =   55
         Tag             =   "41-Other Earning To Date Week 2"
         Top             =   4680
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK3 
         DataField       =   "F7_OTH_EARN_TO_WK3"
         Height          =   285
         Left            =   3240
         TabIndex        =   59
         Tag             =   "41-Other Earning To Date Week 3"
         Top             =   5040
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpOtherEarnToWK4 
         DataField       =   "F7_OTH_EARN_TO_WK4"
         Height          =   285
         Left            =   3240
         TabIndex        =   63
         Tag             =   "41-Other Earning To Date Week 4"
         Top             =   5400
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
         TabIndex        =   52
         Tag             =   "20-Mandatory Overtime Pay Week 1"
         Top             =   4320
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK2 
         DataField       =   "F7_MAND_OVT_PAY_WK2"
         Height          =   285
         Left            =   5040
         TabIndex        =   56
         Tag             =   "20-Mandatory Overtime Pay Week 2"
         Top             =   4680
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK3 
         DataField       =   "F7_MAND_OVT_PAY_WK3"
         Height          =   285
         Left            =   5040
         TabIndex        =   60
         Tag             =   "20-Mandatory Overtime Pay Week 3"
         Top             =   5040
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMandOvtPayWK4 
         DataField       =   "F7_MAND_OVT_PAY_WK4"
         Height          =   285
         Left            =   5040
         TabIndex        =   64
         Tag             =   "20-Mandatory Overtime Pay Week 4"
         Top             =   5400
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK1 
         DataField       =   "F7_VOL_OVT_PAY_WK1"
         Height          =   285
         Left            =   6600
         TabIndex        =   53
         Tag             =   "20-Voluntary Overtime Pay Week 1"
         Top             =   4320
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK2 
         DataField       =   "F7_VOL_OVT_PAY_WK2"
         Height          =   285
         Left            =   6600
         TabIndex        =   57
         Tag             =   "20-Voluntary Overtime Pay Week 2"
         Top             =   4680
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK3 
         DataField       =   "F7_VOL_OVT_PAY_WK3"
         Height          =   285
         Left            =   6600
         TabIndex        =   61
         Tag             =   "20-Voluntary Overtime Pay Week 3"
         Top             =   5040
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVolOvtPayWK4 
         DataField       =   "F7_VOL_OVT_PAY_WK4"
         Height          =   285
         Left            =   6600
         TabIndex        =   65
         Tag             =   "20-Voluntary Overtime Pay Week 4"
         Top             =   5400
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK1 
         DataField       =   "F7_OTH_EARN_1_WK1"
         Height          =   285
         Left            =   1800
         TabIndex        =   70
         Tag             =   "20-Other Earnings 1 - Week 1"
         Top             =   6720
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK2 
         DataField       =   "F7_OTH_EARN_1_WK2"
         Height          =   285
         Left            =   1800
         TabIndex        =   71
         Tag             =   "20-Other Earnings 1 - Week 2"
         Top             =   7080
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK3 
         DataField       =   "F7_OTH_EARN_1_WK3"
         Height          =   285
         Left            =   1800
         TabIndex        =   72
         Tag             =   "20-Other Earnings 1 - Week 3"
         Top             =   7440
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn1WK4 
         DataField       =   "F7_OTH_EARN_1_WK4"
         Height          =   285
         Left            =   1800
         TabIndex        =   73
         Tag             =   "20-Other Earnings 1 - Week 4"
         Top             =   7800
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK1 
         DataField       =   "F7_OTH_EARN_2_WK1"
         Height          =   285
         Left            =   3450
         TabIndex        =   74
         Tag             =   "20-Other Earnings 2 - Week 1"
         Top             =   6720
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK2 
         DataField       =   "F7_OTH_EARN_2_WK2"
         Height          =   285
         Left            =   3450
         TabIndex        =   75
         Tag             =   "20-Other Earnings 2 - Week 2"
         Top             =   7080
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK3 
         DataField       =   "F7_OTH_EARN_2_WK3"
         Height          =   285
         Left            =   3450
         TabIndex        =   76
         Tag             =   "20-Other Earnings 2 - Week 3"
         Top             =   7440
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn2WK4 
         DataField       =   "F7_OTH_EARN_2_WK4"
         Height          =   285
         Left            =   3450
         TabIndex        =   77
         Tag             =   "20-Other Earnings 2 - Week 4"
         Top             =   7800
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK1 
         DataField       =   "F7_OTH_EARN_3_WK1"
         Height          =   285
         Left            =   5040
         TabIndex        =   78
         Tag             =   "20-Other Earnings 3 - Week 1"
         Top             =   6720
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK2 
         DataField       =   "F7_OTH_EARN_3_WK2"
         Height          =   285
         Left            =   5040
         TabIndex        =   79
         Tag             =   "20-Other Earnings 3 - Week 2"
         Top             =   7080
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK3 
         DataField       =   "F7_OTH_EARN_3_WK3"
         Height          =   285
         Left            =   5040
         TabIndex        =   80
         Tag             =   "20-Other Earnings 3 - Week 3"
         Top             =   7440
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn3WK4 
         DataField       =   "F7_OTH_EARN_3_WK4"
         Height          =   285
         Left            =   5040
         TabIndex        =   81
         Tag             =   "20-Other Earnings 3 - Week 4"
         Top             =   7800
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK1 
         DataField       =   "F7_OTH_EARN_4_WK1"
         Height          =   285
         Left            =   6600
         TabIndex        =   82
         Tag             =   "20-Other Earnings 4 - Week 1"
         Top             =   6720
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK2 
         DataField       =   "F7_OTH_EARN_4_WK2"
         Height          =   285
         Left            =   6600
         TabIndex        =   83
         Tag             =   "20-Other Earnings 4 - Week 2"
         Top             =   7080
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK3 
         DataField       =   "F7_OTH_EARN_4_WK3"
         Height          =   285
         Left            =   6600
         TabIndex        =   84
         Tag             =   "20-Other Earnings 4 - Week 3"
         Top             =   7440
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medOthEarn4WK4 
         DataField       =   "F7_OTH_EARN_4_WK4"
         Height          =   285
         Left            =   6600
         TabIndex        =   85
         Tag             =   "20-Other Earnings 4 - Week 4"
         Top             =   7800
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medNormTimeFrom 
         DataField       =   "F7_LAST_DAY_WORK_FTIME"
         Height          =   285
         Left            =   3930
         TabIndex        =   37
         Tag             =   "00-Normal From Working Time"
         Top             =   1200
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
      Begin MSMask.MaskEdBox medNormTimeTo 
         DataField       =   "F7_LAST_DAY_WORK_TTIME"
         Height          =   285
         Left            =   6810
         TabIndex        =   40
         Tag             =   "00-Normal To Working Time"
         Top             =   1200
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
      Begin MSMask.MaskEdBox medHourLastWorked 
         DataField       =   "F7_LAST_WORK_TIME"
         Height          =   285
         Left            =   5760
         TabIndex        =   34
         Tag             =   "00-Time Last Worked"
         Top             =   840
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
         Format          =   "hh:mm"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "| ----------------- Overtime Pay ----------------- |"
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
         Left            =   5040
         TabIndex        =   251
         Top             =   3840
         Width           =   2745
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Net Claim Code of Amount:"
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
         TabIndex        =   206
         Top             =   165
         Width           =   2085
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Federal"
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
         Left            =   2760
         TabIndex        =   205
         Top             =   165
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provincial"
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
         Left            =   4935
         TabIndex        =   204
         Top             =   165
         Width           =   690
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Vacation pay on each cheque?"
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
         TabIndex        =   203
         Top             =   525
         Width           =   2415
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provide percentage"
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
         Left            =   4230
         TabIndex        =   202
         Top             =   525
         Width           =   1395
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Date and hour last worked:"
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
         TabIndex        =   201
         Top             =   885
         Width           =   2100
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Normal working hours on last day worked"
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
         TabIndex        =   200
         Top             =   1245
         Width           =   3090
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5. Actual earnings for last day worked"
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
         TabIndex        =   199
         Top             =   1605
         Width           =   2655
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6. Normal earnings for last day worked"
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
         TabIndex        =   198
         Top             =   1965
         Width           =   2700
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   6360
         TabIndex        =   197
         Top             =   525
         Width           =   120
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   3240
         TabIndex        =   196
         Top             =   1605
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   3240
         TabIndex        =   195
         Top             =   1965
         Width           =   90
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
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
         Left            =   3435
         TabIndex        =   194
         Top             =   1245
         Width           =   345
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   6435
         TabIndex        =   193
         Top             =   1245
         Width           =   195
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7. Advances on wages:"
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
         TabIndex        =   192
         Top             =   2280
         Width           =   1680
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If yes, indicate: "
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
         Left            =   600
         TabIndex        =   191
         Top             =   2580
         Width           =   1110
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8. Other Earnings (Not Regular Wages): "
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
         TabIndex        =   190
         Top             =   2920
         Width           =   2865
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provide the total of additional earnings for each week for the 4 weeks before the accident/illness."
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
         Left            =   480
         TabIndex        =   189
         Top             =   3180
         Width           =   6870
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Left            =   720
         TabIndex        =   188
         Top             =   4035
         Width           =   450
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   2160
         TabIndex        =   187
         Top             =   4035
         Width           =   735
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   3840
         TabIndex        =   186
         Top             =   4035
         Width           =   585
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mandatory"
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
         Left            =   5280
         TabIndex        =   185
         Top             =   4035
         Width           =   750
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voluntary"
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
         Left            =   6840
         TabIndex        =   184
         Top             =   4035
         Width           =   660
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 1"
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
         Left            =   720
         TabIndex        =   183
         Top             =   4365
         Width           =   570
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 2"
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
         Left            =   720
         TabIndex        =   182
         Top             =   4725
         Width           =   570
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 3"
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
         Left            =   720
         TabIndex        =   181
         Top             =   5085
         Width           =   570
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 4"
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
         Left            =   720
         TabIndex        =   180
         Top             =   5445
         Width           =   570
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   179
         Top             =   4365
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   178
         Top             =   4725
         Width           =   90
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   177
         Top             =   5085
         Width           =   90
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   176
         Top             =   5445
         Width           =   90
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   175
         Top             =   4365
         Width           =   90
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   174
         Top             =   4725
         Width           =   90
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   173
         Top             =   5085
         Width           =   90
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   172
         Top             =   5445
         Width           =   90
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEInjF7Sections.frx":0010
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   171
         Top             =   3480
         Width           =   7800
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Use these spaces for any other earnings (indicate Commission, Differentials, Premiums, Bonus, Tips, In Lieu %, etc.)."
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
         Left            =   720
         TabIndex        =   170
         Top             =   5880
         Width           =   8385
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   1680
         TabIndex        =   169
         Top             =   6765
         Width           =   90
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   1680
         TabIndex        =   168
         Top             =   7125
         Width           =   90
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   1680
         TabIndex        =   167
         Top             =   7485
         Width           =   90
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   1680
         TabIndex        =   166
         Top             =   7845
         Width           =   90
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   3320
         TabIndex        =   165
         Top             =   6765
         Width           =   90
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   3320
         TabIndex        =   164
         Top             =   7125
         Width           =   90
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   3320
         TabIndex        =   163
         Top             =   7485
         Width           =   90
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   3320
         TabIndex        =   162
         Top             =   7845
         Width           =   90
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   161
         Top             =   6765
         Width           =   90
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   160
         Top             =   7125
         Width           =   90
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   159
         Top             =   7485
         Width           =   90
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   4920
         TabIndex        =   158
         Top             =   7845
         Width           =   90
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   157
         Top             =   6765
         Width           =   90
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   156
         Top             =   7125
         Width           =   90
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   155
         Top             =   7485
         Width           =   90
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         Left            =   6480
         TabIndex        =   154
         Top             =   7845
         Width           =   90
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Left            =   720
         TabIndex        =   153
         Top             =   6300
         Width           =   450
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 1"
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
         Left            =   720
         TabIndex        =   152
         Top             =   6765
         Width           =   570
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 2"
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
         Left            =   720
         TabIndex        =   151
         Top             =   7125
         Width           =   570
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 3"
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
         Left            =   720
         TabIndex        =   150
         Top             =   7485
         Width           =   570
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Week 4"
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
         Left            =   720
         TabIndex        =   149
         Top             =   7845
         Width           =   570
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Is the worker being paid while he/she recovers?"
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
         Left            =   1920
         TabIndex        =   148
         Top             =   2280
         Width           =   3405
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   2760
         TabIndex        =   147
         Top             =   885
         Width           =   345
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   5280
         TabIndex        =   146
         Top             =   885
         Width           =   345
      End
   End
   Begin VB.Frame frELostTime 
      BorderStyle     =   0  'None
      Caption         =   "E. Lost Time - No Lost Time"
      Height          =   4335
      Left            =   360
      TabIndex        =   123
      Top             =   1800
      Visible         =   0   'False
      Width           =   9615
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   5520
         TabIndex        =   261
         Top             =   2415
         Width           =   2775
         Begin VB.OptionButton optRegMod 
            Caption         =   "modified work"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   7
            Tag             =   "40-Modified Work"
            Top             =   0
            Width           =   1335
         End
         Begin VB.OptionButton optRegMod 
            Caption         =   "regular work"
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
            Index           =   0
            Left            =   0
            TabIndex        =   6
            Tag             =   "40-Regular Work"
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   260
         Top             =   3360
         Width           =   1935
         Begin VB.OptionButton optConfirmedBy 
            Caption         =   "Other"
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
            Left            =   1080
            TabIndex        =   9
            Tag             =   "40-Employee's Premises Yes"
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optConfirmedBy 
            Caption         =   "Myself"
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
            Left            =   0
            TabIndex        =   8
            Tag             =   "40-Employee's Premises Yes"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   360
         TabIndex        =   259
         Top             =   720
         Width           =   7935
         Begin VB.OptionButton optLostTime 
            Caption         =   "Has lost time and/or earnings. (Complete ALL remaining sections)."
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
            Left            =   0
            TabIndex        =   3
            Tag             =   "40-Lost time and/or earnings"
            Top             =   720
            Width           =   5535
         End
         Begin VB.OptionButton optLostTime 
            Caption         =   "Returned to his/her regular job and has not lost any time and/or earnings. (Complete sections G and J)."
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
            Left            =   0
            TabIndex        =   1
            Tag             =   "40-Returned to regular job and has not lost any time and/or earnings"
            Top             =   0
            Width           =   7695
         End
         Begin VB.OptionButton optLostTime 
            Caption         =   "Returned to modified work and has not lost any time and/or earnings. (Complete sections F, G, and J)."
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
            Left            =   0
            TabIndex        =   2
            Tag             =   "40-Returned to modified work and has not lost any time and/or earnings"
            Top             =   360
            Width           =   7575
         End
      End
      Begin VB.TextBox txtConfirmedBy 
         Appearance      =   0  'Flat
         DataField       =   "F7_CONFIRM_BY"
         Height          =   285
         Left            =   8640
         MaxLength       =   5
         TabIndex        =   238
         Top             =   3240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtRegMod 
         Appearance      =   0  'Flat
         DataField       =   "F7_RETURN_REG_MOD"
         Height          =   285
         Left            =   8640
         MaxLength       =   5
         TabIndex        =   237
         Top             =   1920
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtLostTime 
         Appearance      =   0  'Flat
         DataField       =   "F7_RETURNED_TO"
         Height          =   285
         Left            =   8640
         MaxLength       =   5
         TabIndex        =   236
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtConfirmedByName 
         Appearance      =   0  'Flat
         DataField       =   "F7_CONFIRM_NAME"
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
         Left            =   3000
         MaxLength       =   35
         TabIndex        =   10
         Tag             =   "01-Information was confirmed by"
         Top             =   3360
         Width           =   3555
      End
      Begin VB.TextBox txtConfirmPhoneExt 
         Appearance      =   0  'Flat
         DataField       =   "F7_CONFIRM_PHONE_EXT"
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
         Left            =   4920
         MaxLength       =   6
         TabIndex        =   12
         Tag             =   "01-Phone Extension"
         Top             =   3720
         Width           =   795
      End
      Begin INFOHR_Controls.DateLookup dlpRetLostDate 
         DataField       =   "F7_LOST_DATE"
         Height          =   285
         Index           =   0
         Left            =   3840
         TabIndex        =   4
         Tag             =   "41-Date worker first lost time"
         Top             =   1920
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpRetLostDate 
         DataField       =   "F7_RETURN_DATE"
         Height          =   285
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Tag             =   "41-Date worker returned to work (if known)"
         Top             =   2400
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin MSMask.MaskEdBox medConfirmPhone 
         DataField       =   "F7_CONFIRM_PHONE"
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Tag             =   "10-Alternate Telephone Number"
         Top             =   3720
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Please choose one of the following indicators. After the day of accident/awareness of illness, this worker:"
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
         TabIndex        =   130
         Top             =   360
         Width           =   7590
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Provide date worker first lost time"
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
         Left            =   840
         TabIndex        =   129
         Top             =   1965
         Width           =   2460
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Date worker returned to work (if known)"
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
         Left            =   840
         TabIndex        =   128
         Top             =   2445
         Width           =   2940
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "2. This Lost Time - No Lost Time - Modified Work information was confirmed by:"
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
         TabIndex        =   127
         Top             =   3000
         Width           =   5700
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
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
         Left            =   2085
         TabIndex        =   126
         Top             =   3765
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ext."
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
         Left            =   4575
         TabIndex        =   125
         Top             =   3765
         Width           =   270
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   2430
         TabIndex        =   124
         Top             =   3405
         Width           =   420
      End
   End
   Begin VB.Frame frIWorkSchedule 
      BorderStyle     =   0  'None
      Caption         =   "I. Work Schedule (Complete either A, Bor C. Do not include overtime shifts)"
      Height          =   5535
      Left            =   360
      TabIndex        =   208
      Top             =   2160
      Width           =   9615
      Begin VB.TextBox txtWorkSchedule 
         Appearance      =   0  'Flat
         DataField       =   "F7_WORKSCH"
         Height          =   285
         Left            =   5640
         MaxLength       =   5
         TabIndex        =   244
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton optWorkSchedule 
         Caption         =   "(A.) Regular Schedule - Indicate normal work days and hours."
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
         Left            =   120
         TabIndex        =   86
         Tag             =   "40-Regular Work Schedule"
         Top             =   360
         Width           =   4695
      End
      Begin VB.OptionButton optWorkSchedule 
         Caption         =   "(B.) Repeating Rotational Shift Worker - Provide"
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
         Left            =   120
         TabIndex        =   94
         Tag             =   "40-Repeating Rotational Shift Worker"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.OptionButton optWorkSchedule 
         Caption         =   "(C.) Varied or Irregular Work Schedule - Provide the total number of regular hours and shifts for each week for the"
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
         Left            =   120
         TabIndex        =   99
         Tag             =   "40-Varied or Irregular Work Schedule"
         Top             =   2760
         Width           =   8295
      End
      Begin MSMask.MaskEdBox medNoDayOn 
         DataField       =   "F7_NUM_DAYS_ON"
         Height          =   285
         Left            =   1695
         TabIndex        =   95
         Tag             =   "11-Number of Days On"
         Top             =   2040
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
         Left            =   3615
         TabIndex        =   96
         Tag             =   "11-Number of Days Off"
         Top             =   2040
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
         Left            =   5565
         TabIndex        =   97
         Tag             =   "11-Hours per Shift(s)"
         Top             =   2040
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
         Left            =   8100
         TabIndex        =   98
         Tag             =   "11-Number of Weeks in Cycle"
         Top             =   2040
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
         TabIndex        =   100
         Tag             =   "41-Week 1 From Date"
         Top             =   3675
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk1TDate 
         DataField       =   "F7_TWEEK1"
         Height          =   285
         Left            =   2280
         TabIndex        =   101
         Tag             =   "41-Week 1 To Date"
         Top             =   4035
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk2FDate 
         DataField       =   "F7_FWEEK2"
         Height          =   285
         Left            =   4080
         TabIndex        =   102
         Tag             =   "41-Week 2 From Date"
         Top             =   3675
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk2TDate 
         DataField       =   "F7_TWEEK2"
         Height          =   285
         Left            =   4080
         TabIndex        =   103
         Tag             =   "41-Week 2 To Date"
         Top             =   4035
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk3FDate 
         DataField       =   "F7_FWEEK3"
         Height          =   285
         Left            =   5880
         TabIndex        =   104
         Tag             =   "41-Week 3 From Date"
         Top             =   3675
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk3TDate 
         DataField       =   "F7_TWEEK3"
         Height          =   285
         Left            =   5880
         TabIndex        =   105
         Tag             =   "41-Week 3 To Date"
         Top             =   4035
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk4FDate 
         DataField       =   "F7_FWEEK4"
         Height          =   285
         Left            =   7680
         TabIndex        =   106
         Tag             =   "41-Week 4 From Date"
         Top             =   3675
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpWk4TDate 
         DataField       =   "F7_TWEEK4"
         Height          =   285
         Left            =   7680
         TabIndex        =   107
         Tag             =   "41-Week 4 To Date"
         Top             =   4035
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin MSMask.MaskEdBox medTotShiftsWrkWK1 
         DataField       =   "F7_TOT_SHIFT_WEEK1"
         Height          =   285
         Left            =   2595
         TabIndex        =   112
         Tag             =   "11-Total Shifts Worked Week 1"
         Top             =   4800
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   4395
         TabIndex        =   113
         Tag             =   "11-Total Shifts Worked Week 2"
         Top             =   4800
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   6195
         TabIndex        =   114
         Tag             =   "11-Total Shifts Worked Week 3"
         Top             =   4800
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   7995
         TabIndex        =   115
         Tag             =   "11-Total Shifts Worked Week 4"
         Top             =   4800
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   2595
         TabIndex        =   108
         Tag             =   "11-Total Hours Worked Week 1"
         Top             =   4440
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   4395
         TabIndex        =   109
         Tag             =   "11-Total Hours Worked Week 2"
         Top             =   4440
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   6195
         TabIndex        =   110
         Tag             =   "11-Total Hours Worked Week 3"
         Top             =   4440
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   7995
         TabIndex        =   111
         Tag             =   "11-Total Hours Worked Week 4"
         Top             =   4440
         Width           =   1005
         _ExtentX        =   1773
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
         Left            =   3240
         TabIndex        =   89
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
         Left            =   4440
         TabIndex        =   90
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
         Left            =   6840
         TabIndex        =   92
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
         Left            =   5640
         TabIndex        =   91
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
         Left            =   8040
         TabIndex        =   93
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
         Left            =   840
         TabIndex        =   87
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
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "F7_REG_SCHD_MON"
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   88
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
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   2145
         TabIndex        =   228
         Top             =   720
         Width           =   585
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   8115
         TabIndex        =   227
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   5700
         TabIndex        =   226
         Top             =   720
         Width           =   675
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   7020
         TabIndex        =   225
         Top             =   720
         Width           =   435
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   4395
         TabIndex        =   224
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   3315
         TabIndex        =   223
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblWeekDay 
         Alignment       =   2  'Center
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
         Left            =   960
         TabIndex        =   222
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Dates"
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
         Left            =   720
         TabIndex        =   221
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4 weeks prior to the accident/illness. (Do not include overtime hours or shifts here)."
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
         Left            =   3240
         TabIndex        =   220
         Top             =   3015
         Width           =   5835
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF DAYS ON"
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
         Height          =   435
         Index           =   13
         Left            =   720
         TabIndex        =   219
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF DAYS OFF"
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
         Height          =   435
         Index           =   14
         Left            =   2640
         TabIndex        =   218
         Top             =   1920
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "HOURS PER SHIFT(s)"
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
         Height          =   435
         Index           =   15
         Left            =   4560
         TabIndex        =   217
         Top             =   1920
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "NUMBER OF WEEKS IN CYCLE"
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
         Height          =   435
         Index           =   16
         Left            =   6480
         TabIndex        =   216
         Top             =   1920
         Width           =   1605
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "From Dates"
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
         Left            =   720
         TabIndex        =   215
         Top             =   3720
         Width           =   810
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Hours Worked"
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
         Left            =   720
         TabIndex        =   214
         Top             =   4485
         Width           =   1440
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Shifts Worked"
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
         Left            =   720
         TabIndex        =   213
         Top             =   4845
         Width           =   1410
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 1"
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
         Left            =   2880
         TabIndex        =   212
         Top             =   3375
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 2"
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
         Left            =   4680
         TabIndex        =   211
         Top             =   3375
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 3"
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
         Left            =   6480
         TabIndex        =   210
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Week 4"
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
         Left            =   8280
         TabIndex        =   209
         Top             =   3360
         Width           =   570
      End
   End
   Begin VB.Frame frFReturnToWork 
      BorderStyle     =   0  'None
      Caption         =   "F. Return To Work"
      Height          =   5175
      Left            =   360
      TabIndex        =   131
      Top             =   1800
      Visible         =   0   'False
      Width           =   9615
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   258
         Top             =   4320
         Width           =   2055
         Begin VB.OptionButton optResponsible 
            Caption         =   "Myself"
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
            Left            =   0
            TabIndex        =   23
            Tag             =   "40-Responsible for arranging worker's return to work Myself"
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optResponsible 
            Caption         =   "Other"
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
            Left            =   1200
            TabIndex        =   24
            Tag             =   "40-Responsible for arranging worker's return to work Other"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   600
         TabIndex        =   257
         Top             =   3000
         Width           =   2415
         Begin VB.OptionButton optAccDecl 
            Caption         =   "Accepted"
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
            Left            =   0
            TabIndex        =   19
            Tag             =   "40-Return to Work Offer, Accepted"
            Top             =   0
            Width           =   1095
         End
         Begin VB.OptionButton optAccDecl 
            Caption         =   "Declined"
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
            Left            =   1200
            TabIndex        =   20
            Tag             =   "40-Return to Work Offer, Declined"
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   256
         Top             =   2040
         Width           =   1455
         Begin VB.OptionButton optOffered 
            Caption         =   "Yes"
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
            Left            =   0
            TabIndex        =   17
            Tag             =   "40-Modified work been offered, Yes"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optOffered 
            Caption         =   "No"
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
            Left            =   720
            TabIndex        =   18
            Tag             =   "40-Modified work been offered, No"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         TabIndex        =   255
         Top             =   1320
         Width           =   1455
         Begin VB.OptionButton optDiscussed 
            Caption         =   "Yes"
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
            Left            =   0
            TabIndex        =   15
            Tag             =   "40-Modified work been discussed, Yes"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optDiscussed 
            Caption         =   "No"
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
            Left            =   720
            TabIndex        =   16
            Tag             =   "40-Modified work been discussed, No"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         TabIndex        =   254
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton optLimitations 
            Caption         =   "Yes"
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
            Left            =   120
            TabIndex        =   13
            Tag             =   "40-Provided with work limitations, Yes"
            Top             =   0
            Width           =   615
         End
         Begin VB.OptionButton optLimitations 
            Caption         =   "No"
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
            Left            =   840
            TabIndex        =   14
            Tag             =   "40-Provided with work limitations, No"
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.TextBox txtResponsible 
         Appearance      =   0  'Flat
         DataField       =   "F7_RESPONSIBLE"
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   243
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAccDecl 
         Appearance      =   0  'Flat
         DataField       =   "F7_ACCEPT_DECLINE"
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   242
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtOffered 
         Appearance      =   0  'Flat
         DataField       =   "F7_OFFERED"
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   241
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDiscussed 
         Appearance      =   0  'Flat
         DataField       =   "F7_DISCUSSED"
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   240
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtLimitations 
         Appearance      =   0  'Flat
         DataField       =   "F7_LIMITATION"
         Height          =   285
         Left            =   4560
         MaxLength       =   5
         TabIndex        =   239
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtResponsiblePhoneExt 
         Appearance      =   0  'Flat
         DataField       =   "F7_RESPONS_PHONE_EXT"
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
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   27
         Tag             =   "01-Phone Extension"
         Top             =   4680
         Width           =   795
      End
      Begin VB.TextBox txtResponsibleName 
         Appearance      =   0  'Flat
         DataField       =   "F7_RESPONS_NAME"
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
         Left            =   3120
         MaxLength       =   35
         TabIndex        =   25
         Tag             =   "01-Responsible for arranging worker's return to work"
         Top             =   4320
         Width           =   3555
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   330
         Left            =   8640
         TabIndex        =   22
         Top             =   3292
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSMask.MaskEdBox medResponsiblePhone 
         DataField       =   "F7_RESPONS_PHONE"
         Height          =   285
         Left            =   3120
         TabIndex        =   26
         Tag             =   "10-Telephone Number"
         Top             =   4680
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
      Begin VB.CheckBox chkWrittenOfferAttached 
         Caption         =   "If Declined please attach a copy of the written offer given to the worker."
         DataField       =   "F7_DECLINE_ATTACHED"
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
         Left            =   1680
         TabIndex        =   21
         Tag             =   "Written Offer attachment"
         Top             =   3360
         Width           =   5445
      End
      Begin VB.Label lblImport1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Written Offer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   7275
         TabIndex        =   132
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Have you been provided with work limitations for this worker's injury?"
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
         TabIndex        =   140
         Top             =   360
         Width           =   5010
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Has modified work been discussed with this worker?"
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
         TabIndex        =   139
         Top             =   1080
         Width           =   3870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Has modified work been offered to this worker?"
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
         TabIndex        =   138
         Top             =   1800
         Width           =   3510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If yes, was it "
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
         Left            =   360
         TabIndex        =   137
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   2610
         TabIndex        =   136
         Top             =   4365
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Who is responsible for arranging worker's return to work?"
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
         TabIndex        =   135
         Top             =   3960
         Width           =   4185
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
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
         Left            =   2265
         TabIndex        =   134
         Top             =   4725
         Width           =   765
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ext."
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
         Left            =   4695
         TabIndex        =   133
         Top             =   4725
         Width           =   270
      End
      Begin VB.Image imgNoSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8280
         Picture         =   "frmEInjF7Sections.frx":00C8
         Top             =   3330
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgSec 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   8280
         Picture         =   "frmEInjF7Sections.frx":0212
         Top             =   3337
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.TextBox txtIncidentDate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   252
      TabStop         =   0   'False
      Tag             =   "01-Incident #"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LUSER"
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
      Left            =   8640
      MaxLength       =   25
      TabIndex        =   235
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LTIME"
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
      Left            =   6960
      MaxLength       =   25
      TabIndex        =   234
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LDATE"
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
      Left            =   5160
      MaxLength       =   25
      TabIndex        =   233
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   120
      Top             =   10080
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Caption         =   "Ado2"
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
      Height          =   540
      Left            =   0
      TabIndex        =   121
      Top             =   10020
      Width           =   10410
      _Version        =   65536
      _ExtentX        =   18362
      _ExtentY        =   952
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.CommandButton cmdRefresh 
         Appearance      =   0  'Flat
         Caption         =   "&Refresh Data"
         Height          =   375
         Left            =   7560
         TabIndex        =   232
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   4830
         TabIndex        =   230
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Save Changes"
         Height          =   375
         Left            =   3120
         TabIndex        =   231
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSComctlLib.TabStrip tbF7Sections 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   15901
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "E. Lost Time - No Lost Time"
            Key             =   "tbkeyELostTime"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "F. Return to Work"
            Key             =   "tbkeyFReturnWork"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "H. Additional Wage Information"
            Key             =   "tbkeyHAddWageInfo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "I. Work Schedule"
            Key             =   "tbkeyIWorkSch"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "J. Form 7 Filled By"
            Key             =   "tbkeyJFilledBy"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "K. Additional Information"
            Key             =   "tbkeyKAddInfo"
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
   Begin VB.TextBox txtIncidentNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   122
      TabStop         =   0   'False
      Tag             =   "01-Incident #"
      Top             =   600
      Width           =   870
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   117
      Top             =   0
      Width           =   10410
      _Version        =   65536
      _ExtentX        =   18362
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   255
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
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   120
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "lblEEName"
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
         Left            =   2880
         TabIndex        =   118
         Top             =   135
         Width           =   1185
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
         Left            =   1200
         TabIndex        =   119
         Top             =   135
         Width           =   1245
      End
   End
   Begin VB.Label Label73 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Date"
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
      Left            =   3000
      TabIndex        =   253
      Top             =   645
      Width           =   960
   End
   Begin VB.Label lblIncident 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   240
      TabIndex        =   207
      Top             =   645
      Width           =   1170
   End
End
Attribute VB_Name = "frmEInjF7Sections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbTERM_Seq
Dim xUpdateable
Dim fglbJobList As String
Dim rsDATA As New ADODB.Recordset

Private Function chkHSInjuryF7Sections()
Dim X%, Y%
Dim Msg$, Response%
Dim flgHoursEntered As Boolean
Dim errMsg As Boolean
Dim Part1, Part2

chkHSInjuryF7Sections = False

On Error GoTo chkHSInjuryF7_Err

'Section E: Lost Time - No Lost Time
If optLostTime(0).Value = False And optLostTime(1).Value = False And optLostTime(2).Value = False Then
    tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
    MsgBox "One of the '1. After the day of accident/awareness of illness, this worker' indicator selection is required.", vbExclamation
    'optLostTime(0).SetFocus
    'Exit Function
End If

If optLostTime(2).Value = True Then
    If Len(Trim(dlpRetLostDate(0).Text)) > 0 Then
        If Not IsDate(dlpRetLostDate(0).Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
            MsgBox "Invalid 'Provide date worker first lost time' Date.", vbExclamation
            dlpRetLostDate(0).SetFocus
            Exit Function
        End If
    Else
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
        MsgBox "'Provide date worker first lost time' is required if 'Has lost time and/or earnings' is checked.", vbExclamation
        'dlpRetLostDate(0).SetFocus
        'Exit Function
    End If
End If

If Len(Trim(dlpRetLostDate(1).Text)) > 0 Then
    If Not IsDate(dlpRetLostDate(1).Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
        MsgBox "Invalid 'Date worker returned to work' Date.", vbExclamation
        dlpRetLostDate(1).SetFocus
        Exit Function
    Else
        If optRegMod(0).Value = False And optRegMod(1).Value = False Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
            MsgBox "Either 'regular work' or 'modified work' is required to be checked if 'Date worker returned to work' is entered.", vbExclamation
            'optRegMod(0).SetFocus
            'Exit Function
        End If
    End If
End If

If optConfirmedBy(0).Value = False And optConfirmedBy(1).Value = False Then
    tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
    MsgBox "One of the '2. This Lost Time-No Lost Time-Modified Worker information was confirmed by' indicator selection is required.", vbExclamation
    'optConfirmedBy(0).SetFocus
    'Exit Function
End If

If optConfirmedBy(1).Value = True Then
    If Len(Trim(txtConfirmedByName.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
        MsgBox "Name is required if 'Other' is checked.", vbExclamation
        txtConfirmedByName.SetFocus
        Exit Function
    End If
End If

'Lambton said if Myself is selected then Telephone is not mandatory
If optConfirmedBy(1).Value = True Then
    If Len(Trim(medConfirmPhone.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(1)
        MsgBox "Telephone is required", vbExclamation
        medConfirmPhone.SetFocus
        Exit Function
    End If
End If

'Section F: Return to Work
If optLostTime(0).Value = False Then     'Section E #1 - first checkbox selected - only requires Sec. G and J to be filled
    If optLimitations(0).Value = False And optLimitations(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "One of the '1. Have you been provided with work limitations for this workers injury?' indicator selection is required.", vbExclamation
        'optLimitations(0).SetFocus
        'Exit Function
    End If
    
    If optDiscussed(0).Value = False And optDiscussed(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "One of the '2. Has modified work been discussed with this worker?' indicator selection is required.", vbExclamation
        'optDiscussed(0).SetFocus
        'Exit Function
    End If
    
    If optOffered(0).Value = False And optOffered(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "One of the '3. Has modified work been offered to this worker?' indicator selection is required.", vbExclamation
        'optOffered(0).SetFocus
        'Exit Function
    End If
    
    If optOffered(0).Value = True And optAccDecl(0).Value = False And optAccDecl(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "3. Worker has been offered modified work, 'Accepted' or 'Declined' is required.", vbExclamation
        'optAccDecl(0).SetFocus
        'Exit Function
    End If
    
    If chkWrittenOfferAttached.Value = 1 And imgNoSec.Visible = True Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "'Copy of a Written Offer given to worker' is selected, but no written offer attached.", vbExclamation
        chkWrittenOfferAttached.SetFocus
        Exit Function
    ElseIf chkWrittenOfferAttached.Value <> 1 And imgNoSec.Visible = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "'Copy of a Written Offer given to worker' not is selected, but written offer is attached.", vbExclamation
        chkWrittenOfferAttached.SetFocus
        Exit Function
    End If
    
    If optResponsible(0).Value = False And optResponsible(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "One of the '4. Who is responsible for arranging worker's return to work?' indicator selection is required.", vbExclamation
        'optResponsible(0).SetFocus
        'Exit Function
    End If
    
    If optResponsible(1).Value = True Then
        If Len(Trim(txtResponsibleName.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
            MsgBox "Name is required if 'Other' is checked.", vbExclamation
            txtResponsibleName.SetFocus
            Exit Function
        End If
    End If
    
    'Lambton said if Myself is selected then Telephone is not mandatory
    If optResponsible(1).Value = True Then
        If Len(Trim(medResponsiblePhone.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
            MsgBox "Telephone is required", vbExclamation
            medResponsiblePhone.SetFocus
            Exit Function
        End If
    End If
Else
    'By any chance document attachment selected, it has to be attached.
    If chkWrittenOfferAttached.Value = 1 And imgNoSec.Visible = True Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "'Copy of a Written Offer given to worker' is selected, but no written offer attached.", vbExclamation
        chkWrittenOfferAttached.SetFocus
        Exit Function
    ElseIf chkWrittenOfferAttached.Value <> 1 And imgNoSec.Visible = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(2)
        MsgBox "'Copy of a Written Offer given to worker' not is selected, but written offer is attached.", vbExclamation
        chkWrittenOfferAttached.SetFocus
        Exit Function
    End If
End If

'Section H: Additional Wage Information
If optLostTime(2).Value = True Then     'Section E #1 - third checkbox selected - only requires all sections to be filled
    If Len(Trim(txtFedCodeAmt.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Federal' for '1. Net Claim Code of Amount' is required.", vbExclamation
        'txtFedCodeAmt.SetFocus
        'Exit Function
    End If
    
    If Len(Trim(txtProvCodeAmt.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Provincial' for '1. Net Claim Code of Amount' is required.", vbExclamation
        'txtProvCodeAmt.SetFocus
        'Exit Function
    End If
    
    If optVacPerctYN(0).Value = False And optVacPerctYN(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "One of the 'Vacation Pay - on each cheque' indicator selection is required.", vbExclamation
        'optVacPerctYN(0).SetFocus
        'Exit Function
    End If
    
    If optVacPerctYN(0).Value = True Then
        If Len(Trim(medVacPerct.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Provide Percentage' for 'Vacation Pay - on each cheque' is required.", vbExclamation
            medVacPerct.SetFocus
            Exit Function
        ElseIf Not IsNumeric(medVacPerct.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "Invalid 'Provide Percentage' for 'Vacation Pay - on each cheque'.", vbExclamation
            medVacPerct.SetFocus
            Exit Function
        ElseIf Val(medVacPerct.Text) <= 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Provide Percentage' for 'Vacation Pay - on each cheque' cannot be less or equal to 0.", vbExclamation
            medVacPerct.SetFocus
            Exit Function
        End If
    End If
    
    '3. Date and hour last worked
    If Len(Trim(dlpDateLastWork.Text)) = 0 Or Len(Trim(medHourLastWorked.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Date and hour last worked' is required.", vbExclamation
        If Len(Trim(dlpDateLastWork.Text)) = 0 Then
            'dlpDateLastWork.SetFocus
        Else
            'medHourLastWorked.SetFocus
        End If
        'Exit Function
    ElseIf Not IsDate(dlpDateLastWork.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "Invalid 'Date' for 'Date and hour last worked'.", vbExclamation
        dlpDateLastWork.SetFocus
        Exit Function
    End If
    
    If Len(medHourLastWorked.Text) = 5 Then
        Part1 = Left(medHourLastWorked, 2)
        Part2 = Right(medHourLastWorked, 2)
        If Not Left(Part1, 2) = "__" Or Not Right(Part2, 2) = "__" Then
            If Not IsNumeric(Part1) Or Not IsNumeric(Part2) Then
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
                MsgBox "Invalid 'Time' for 'Date and hour last worked'.", vbExclamation
                medHourLastWorked.SetFocus
                Exit Function
            End If
            If CInt(Part1) > 12 Or CInt(Part2) > 59 Then
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
                MsgBox "Invalid 'Time' for 'Date and hour last worked'.", vbExclamation
                medHourLastWorked.SetFocus
                Exit Function
            End If
        End If
    ElseIf Len(medHourLastWorked.Text) <> 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "Invalid 'Time' for 'Date and hour last worked'.", vbExclamation
        medHourLastWorked.SetFocus
        Exit Function
    End If
    
    If optTimeLastWorkAP(0).Value = False And optTimeLastWorkAP(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "One of the 'Time' indicator selection for 'Date and hour last worked' is required.", vbExclamation
        'optTimeLastWorkAP(0).SetFocus
        'Exit Function
    End If
    
    '4. Normal working hours on last day worked
    If Len(medNormTimeFrom.Text) = 5 Then
        Part1 = Left(medNormTimeFrom, 2)
        Part2 = Right(medNormTimeFrom, 2)
        If Not Left(Part1, 2) = "__" Or Not Right(Part2, 2) = "__" Then
            If Not IsNumeric(Part1) Or Not IsNumeric(Part2) Then
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
                MsgBox "Invalid 'From' time for '4. Normal working hours on last day worked'.", vbExclamation
                medNormTimeFrom.SetFocus
                Exit Function
            End If
            If CInt(Part1) > 12 Or CInt(Part2) > 59 Then
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
                MsgBox "Invalid 'From' time for '4. Normal working hours on last day worked'.", vbExclamation
                medNormTimeFrom.SetFocus
                Exit Function
            End If
        End If
    ElseIf Len(medNormTimeFrom.Text) <> 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "Invalid 'From' time for '4. Normal working hours on last day worked'.", vbExclamation
        medNormTimeFrom.SetFocus
        Exit Function
    End If
    
    If optNormFTimeLastWorkAP(0).Value = False And optNormFTimeLastWorkAP(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "One of the 'From' time indicator selection for '4. Normal working hours on last day worked' is required.", vbExclamation
        'optNormFTimeLastWorkAP(0).SetFocus
        'Exit Function
    End If
    
    If Len(medNormTimeTo.Text) = 5 Then
        Part1 = Left(medNormTimeTo, 2)
        Part2 = Right(medNormTimeTo, 2)
        If Not Left(Part1, 2) = "__" Or Not Right(Part2, 2) = "__" Then
            If Not IsNumeric(Part1) Or Not IsNumeric(Part2) Then
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
                MsgBox "Invalid 'To' time for '4. Normal working hours on last day worked'.", vbExclamation
                medNormTimeTo.SetFocus
                Exit Function
            End If
            If CInt(Part1) > 12 Or CInt(Part2) > 59 Then
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
                MsgBox "Invalid 'To' time for '4. Normal working hours on last day worked'.", vbExclamation
                medNormTimeTo.SetFocus
                Exit Function
            End If
        End If
    ElseIf Len(medNormTimeTo.Text) <> 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "Invalid 'To' time for '4. Normal working hours on last day worked'.", vbExclamation
        medNormTimeTo.SetFocus
        Exit Function
    End If
    
    If optNormTTimeLastWorkAP(0).Value = False And optNormTTimeLastWorkAP(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "One of the 'To' time indicator selection for '4. Normal working hours on last day worked' is required.", vbExclamation
        'optNormTTimeLastWorkAP(0).SetFocus
        'Exit Function
    End If
    
    '5. Actual earnings for last day worked
    If Len(Trim(medLastActualEarnings.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'5. Actual earnings for last day worked' is required.", vbExclamation
        'medLastActualEarnings.SetFocus
        'Exit Function
    ElseIf Not IsNumeric(medLastActualEarnings.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "Invalid '5. Actual earnings for last day worked'.", vbExclamation
        medLastActualEarnings.SetFocus
        Exit Function
    ElseIf Val(medLastActualEarnings.Text) < 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'5. Actual earnings for last day worked' cannot be less than 0.", vbExclamation
        medLastActualEarnings.SetFocus
        Exit Function
    End If
    
    '6. Normal earnings for last day worked
    If Len(Trim(medLastNormalEarnings.Text)) = 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'6. Normal earnings for last day worked' is required.", vbExclamation
        'medLastNormalEarnings.SetFocus
        'Exit Function
    ElseIf Not IsNumeric(medLastNormalEarnings.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "Invalid '6. Normal earnings for last day worked'.", vbExclamation
        medLastNormalEarnings.SetFocus
        Exit Function
    ElseIf Val(medLastNormalEarnings.Text) < 0 Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'6. Normal earnings for last day worked' cannot be less than 0.", vbExclamation
        medLastNormalEarnings.SetFocus
        Exit Function
    End If
    
    '7. Advances on wages:
    If optBeingPaidYN(0).Value = False And optBeingPaidYN(1).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "One of the 'Is the worker being paid while he/she recovers?' indicator selection for '7. Advances on wages' is required.", vbExclamation
        'optBeingPaidYN(0).SetFocus
        'Exit Function
    End If
    
    If optFullRegOther(1).Value = True Then
        If Len(Trim(txtBeingPaidOther.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "Specify 'Other' when 'Other' is checked.", vbExclamation
            txtBeingPaidOther.SetFocus
            Exit Function
        End If
    End If
    
    '8. Other Earnings (Not Regular Wages)
    If IsDate(dlpOtherEarnFromWK1.Text) And Not IsDate(dlpOtherEarnToWK1.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Other Earnings - Week 1 To Date' is required when 'Week 1 From Date' is entered.", vbExclamation
        'dlpOtherEarnToWK1.SetFocus
        'Exit Function
    End If
    If IsDate(dlpOtherEarnFromWK2.Text) And Not IsDate(dlpOtherEarnToWK2.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Other Earnings - Week 2 To Date' is required when 'Week 2 From Date' is entered.", vbExclamation
        'dlpOtherEarnToWK2.SetFocus
        'Exit Function
    End If
    If IsDate(dlpOtherEarnFromWK3.Text) And Not IsDate(dlpOtherEarnToWK3.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Other Earnings - Week 3 To Date' is required when 'Week 3 From Date' is entered.", vbExclamation
        'dlpOtherEarnToWK3.SetFocus
        'Exit Function
    End If
    If IsDate(dlpOtherEarnFromWK4.Text) And Not IsDate(dlpOtherEarnToWK4.Text) Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
        MsgBox "'Other Earnings - Week 4 To Date' is required when 'Week 4 From Date' is entered.", vbExclamation
        'dlpOtherEarnToWK4.SetFocus
        'Exit Function
    End If
    
    'For X = 1 To 4
    '    errMsg = False
    '    If IsDate("dlpOtherEarnFromWK" & X) And Not IsDate("dlpOtherEarnToWK" & X) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Other Earnings - Week " & X & " To Date' is required when 'Week " & X & " From Date' is entered.", vbExclamation
    '        errMsg = True
    '    End If
    '
    '    If errMsg Then
    '        Select Case X
    '            Case 1: dlpOtherEarnToWK1.SetFocus
    '            Case 2: dlpOtherEarnToWK2.SetFocus
    '            Case 3: dlpOtherEarnToWK3.SetFocus
    '            Case 4: dlpOtherEarnToWK4.SetFocus
    '        End Select
    '        Exit Function
    '    End If
    'Next
    
    If IsDate(dlpOtherEarnFromWK1.Text) And IsDate(dlpOtherEarnToWK1.Text) Then
        If CVDate(dlpOtherEarnFromWK1.Text) > CVDate(dlpOtherEarnToWK1.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Other Earnings - Week 1 From Date' cannot be greater than 'Other Earnings - Week 1 To Date'.", vbExclamation
            dlpOtherEarnFromWK1.SetFocus
            Exit Function
        End If
    End If
    If IsDate(dlpOtherEarnFromWK2.Text) And IsDate(dlpOtherEarnToWK2.Text) Then
        If CVDate(dlpOtherEarnFromWK2.Text) > CVDate(dlpOtherEarnToWK2.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Other Earnings - Week 2 From Date' cannot be greater than 'Other Earnings - Week 2 To Date'.", vbExclamation
            dlpOtherEarnFromWK2.SetFocus
            Exit Function
        End If
    End If
    If IsDate(dlpOtherEarnFromWK3.Text) And IsDate(dlpOtherEarnToWK3.Text) Then
        If CVDate(dlpOtherEarnFromWK3.Text) > CVDate(dlpOtherEarnToWK3.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Other Earnings - Week 3 From Date' cannot be greater than 'Other Earnings - Week 3 To Date'.", vbExclamation
            dlpOtherEarnFromWK3.SetFocus
            Exit Function
        End If
    End If
    If IsDate(dlpOtherEarnFromWK4.Text) And IsDate(dlpOtherEarnToWK4.Text) Then
        If CVDate(dlpOtherEarnFromWK4.Text) > CVDate(dlpOtherEarnToWK4.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Other Earnings - Week 4 From Date' cannot be greater than 'Other Earnings - Week 4 To Date'.", vbExclamation
            dlpOtherEarnFromWK4.SetFocus
            Exit Function
        End If
    End If
    
    'For X = 1 To 4
    '    If IsDate("dlpOtherEarnFromWK" & X) And IsDate("dlpOtherEarnToWK" & X) Then
    '        errMsg = False
    '        If CVDate("dlpOtherEarnFromWK" & X) > CVDate("dlpOtherEarnToWK" & X) Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '            MsgBox "'Other Earnings - Week " & X & " From Date' cannot be greater than 'Other Earnings - Week " & X & " To Date'.", vbExclamation
    '            errMsg = True
    '        End If
    '        If errMsg Then
    '            Select Case X
    '                Case 1: dlpOtherEarnFromWK1.SetFocus
    '                Case 2: dlpOtherEarnFromWK2.SetFocus
    '                Case 3: dlpOtherEarnFromWK3.SetFocus
    '                Case 4: dlpOtherEarnFromWK4.SetFocus
    '            End Select
    '            Exit Function
    '        End If
    '    End If
    'Next
    
    'This is not mandatory
    'Week 1 - 4: Other Earning 1
    If IsDate(dlpOtherEarnFromWK1.Text) And IsDate(dlpOtherEarnToWK1.Text) Then
    '    If Trim(comOtherEarnings1.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 1 cannot be blank.", vbExclamation
    '        comOtherEarnings1.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn1WK1.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 1 for Week 1 cannot be blank.", vbExclamation
    '        medOthEarn1WK1.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn1WK1.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 1 for Week 1.", vbExclamation
    '        medOthEarn1WK1.SetFocus
    '        Exit Function
        If Val(medOthEarn1WK1.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 1 for Week 1 cannot be less than 0.", vbExclamation
            medOthEarn1WK1.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK2.Text) And IsDate(dlpOtherEarnToWK2.Text) Then
    '    If Trim(comOtherEarnings1.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 1 cannot be blank.", vbExclamation
    '        comOtherEarnings1.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn1WK2.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 1 for Week 2 cannot be blank.", vbExclamation
    '        medOthEarn1WK2.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn1WK2.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 1 for Week 2.", vbExclamation
    '        medOthEarn1WK2.SetFocus
    '        Exit Function
        If Val(medOthEarn1WK2.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 1 for Week 2 cannot be less than 0.", vbExclamation
            medOthEarn1WK2.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK3.Text) And IsDate(dlpOtherEarnToWK3.Text) Then
    '    If Trim(comOtherEarnings1.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 1 cannot be blank.", vbExclamation
    '        comOtherEarnings1.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn1WK3.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 1 for Week 3 cannot be blank.", vbExclamation
    '        medOthEarn1WK3.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn1WK3.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 1 for Week 3.", vbExclamation
    '        medOthEarn1WK3.SetFocus
    '        Exit Function
        If Val(medOthEarn1WK3.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 1 for Week 3 cannot be less than 0.", vbExclamation
            medOthEarn1WK3.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK4.Text) And IsDate(dlpOtherEarnToWK4.Text) Then
    '    If Trim(comOtherEarnings1.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 1 cannot be blank.", vbExclamation
    '        comOtherEarnings1.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn1WK4.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 1 for Week 4 cannot be blank.", vbExclamation
    '        medOthEarn1WK4.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn1WK4.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 1 for Week 4.", vbExclamation
    '        medOthEarn1WK4.SetFocus
    '        Exit Function
        If Val(medOthEarn1WK4.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 1 for Week 4 cannot be less than 0.", vbExclamation
            medOthEarn1WK4.SetFocus
            Exit Function
        End If
    End If
    
    
    'Week 1 - 4: Other Earning 2
    If IsDate(dlpOtherEarnFromWK1.Text) And IsDate(dlpOtherEarnToWK1.Text) Then
    '    If Trim(comOtherEarnings2.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 2 cannot be blank.", vbExclamation
    '        comOtherEarnings2.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn2WK1.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 2 for Week 1 cannot be blank.", vbExclamation
    '        medOthEarn2WK1.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn2WK1.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 2 for Week 1.", vbExclamation
    '        medOthEarn2WK1.SetFocus
    '        Exit Function
        If Val(medOthEarn2WK1.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 2 for Week 1 cannot be less than 0.", vbExclamation
            medOthEarn2WK1.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK2.Text) And IsDate(dlpOtherEarnToWK2.Text) Then
    '    If Trim(comOtherEarnings2.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 2 cannot be blank.", vbExclamation
    '        comOtherEarnings2.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn2WK2.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 2 for Week 2 cannot be blank.", vbExclamation
    '        medOthEarn2WK2.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn2WK2.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 2 for Week 2.", vbExclamation
    '        medOthEarn2WK2.SetFocus
    '        Exit Function
        If Val(medOthEarn2WK2.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 2 for Week 2 cannot be less than 0.", vbExclamation
            medOthEarn2WK2.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK3.Text) And IsDate(dlpOtherEarnToWK3.Text) Then
    '    If Trim(comOtherEarnings2.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 2 cannot be blank.", vbExclamation
    '        comOtherEarnings2.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn2WK3.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 2 for Week 3 cannot be blank.", vbExclamation
    '        medOthEarn2WK3.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn2WK3.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 2 for Week 3.", vbExclamation
    '        medOthEarn2WK3.SetFocus
    '        Exit Function
        If Val(medOthEarn2WK3.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 2 for Week 3 cannot be less than 0.", vbExclamation
            medOthEarn2WK3.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK4.Text) And IsDate(dlpOtherEarnToWK4.Text) Then
    '    If Trim(comOtherEarnings2.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 2 cannot be blank.", vbExclamation
    '        comOtherEarnings2.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn2WK4.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 2 for Week 4 cannot be blank.", vbExclamation
    '        medOthEarn2WK4.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn2WK4.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 2 for Week 4.", vbExclamation
    '        medOthEarn2WK4.SetFocus
    '        Exit Function
        If Val(medOthEarn2WK4.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 2 for Week 4 cannot be less than 0.", vbExclamation
            medOthEarn2WK4.SetFocus
            Exit Function
        End If
    End If
    
    'Week 1 - 4: Other Earning 3
    If IsDate(dlpOtherEarnFromWK1.Text) And IsDate(dlpOtherEarnToWK1.Text) Then
    '    If Trim(comOtherEarnings3.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 3 cannot be blank.", vbExclamation
    '        comOtherEarnings3.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn3WK1.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 3 for Week 1 cannot be blank.", vbExclamation
    '        medOthEarn3WK1.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn3WK1.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 3 for Week 1.", vbExclamation
    '        medOthEarn3WK1.SetFocus
    '        Exit Function
        If Val(medOthEarn3WK1.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 3 for Week 1 cannot be less than 0.", vbExclamation
            medOthEarn3WK1.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK2.Text) And IsDate(dlpOtherEarnToWK2.Text) Then
    '    If Trim(comOtherEarnings3.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 3 cannot be blank.", vbExclamation
    '        comOtherEarnings3.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn3WK2.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 3 for Week 2 cannot be blank.", vbExclamation
    '        medOthEarn3WK2.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn3WK2.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 3 for Week 2.", vbExclamation
    '        medOthEarn3WK2.SetFocus
    '        Exit Function
        If Val(medOthEarn3WK2.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 3 for Week 2 cannot be less than 0.", vbExclamation
            medOthEarn3WK2.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK3.Text) And IsDate(dlpOtherEarnToWK3.Text) Then
    '    If Trim(comOtherEarnings3.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 3 cannot be blank.", vbExclamation
    '        comOtherEarnings3.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn3WK3.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 3 for Week 3 cannot be blank.", vbExclamation
    '        medOthEarn3WK3.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn3WK3.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 3 for Week 3.", vbExclamation
    '        medOthEarn3WK3.SetFocus
    '        Exit Function
        If Val(medOthEarn3WK3.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 3 for Week 3 cannot be less than 0.", vbExclamation
            medOthEarn3WK3.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK4.Text) And IsDate(dlpOtherEarnToWK4.Text) Then
    '    If Trim(comOtherEarnings3.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 3 cannot be blank.", vbExclamation
    '        comOtherEarnings3.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn3WK4.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 3 for Week 4 cannot be blank.", vbExclamation
    '        medOthEarn3WK4.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn3WK4.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 3 for Week 4.", vbExclamation
    '        medOthEarn3WK4.SetFocus
    '        Exit Function
        If Val(medOthEarn3WK4.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 3 for Week 4 cannot be less than 0.", vbExclamation
            medOthEarn3WK4.SetFocus
            Exit Function
        End If
    End If
    
    'Week 1 - 4: Other Earning 4
    If IsDate(dlpOtherEarnFromWK1.Text) And IsDate(dlpOtherEarnToWK1.Text) Then
    '    If Trim(comOtherEarnings4.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 4 cannot be blank.", vbExclamation
    '        comOtherEarnings4.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn4WK1.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 4 for Week 1 cannot be blank.", vbExclamation
    '        medOthEarn4WK1.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn4WK1.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 4 for Week 1.", vbExclamation
    '        medOthEarn4WK1.SetFocus
    '        Exit Function
        If Val(medOthEarn4WK1.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 4 for Week 1 cannot be less than 0.", vbExclamation
            medOthEarn4WK1.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK2.Text) And IsDate(dlpOtherEarnToWK2.Text) Then
    '    If Trim(comOtherEarnings4.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 4 cannot be blank.", vbExclamation
    '        comOtherEarnings4.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn4WK2.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 4 for Week 2 cannot be blank.", vbExclamation
    '        medOthEarn4WK2.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn4WK2.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 4 for Week 2.", vbExclamation
    '        medOthEarn4WK2.SetFocus
    '        Exit Function
        If Val(medOthEarn4WK2.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 4 for Week 2 cannot be less than 0.", vbExclamation
            medOthEarn4WK2.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK3.Text) And IsDate(dlpOtherEarnToWK3.Text) Then
    '    If Trim(comOtherEarnings4.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 4 cannot be blank.", vbExclamation
    '        comOtherEarnings4.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn4WK3.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 4 for Week 3 cannot be blank.", vbExclamation
    '        medOthEarn4WK3.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn4WK3.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 4 for Week 3.", vbExclamation
    '        medOthEarn4WK3.SetFocus
    '        Exit Function
        If Val(medOthEarn4WK3.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 4 for Week 3 cannot be less than 0.", vbExclamation
            medOthEarn4WK3.SetFocus
            Exit Function
        End If
    End If
    
    If IsDate(dlpOtherEarnFromWK4.Text) And IsDate(dlpOtherEarnToWK4.Text) Then
    '    If Trim(comOtherEarnings4.Text) = "" Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' type for column 4 cannot be blank.", vbExclamation
    '        comOtherEarnings4.SetFocus
    '        Exit Function
    '    ElseIf Len(medOthEarn4WK4.Text) = 0 Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "'Any Other Earnings' for column 4 for Week 4 cannot be blank.", vbExclamation
    '        medOthEarn4WK4.SetFocus
    '        Exit Function
    '    ElseIf Not IsNumeric(medOthEarn4WK4.Text) Then
    '        tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '        MsgBox "Invalid 'Any Other Earnings' for column 4 for Week 4.", vbExclamation
    '        medOthEarn4WK4.SetFocus
    '        Exit Function
        If Val(medOthEarn4WK4.Text) < 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
            MsgBox "'Any Other Earnings' for column 4 for Week 4 cannot be less than 0.", vbExclamation
            medOthEarn4WK4.SetFocus
            Exit Function
        End If
    End If
    
    
    'For Y = 1 To 4  'Weeks
    '    If IsDate("dlpOtherEarnFromWK" & Y) And IsDate("dlpOtherEarnToWK" & Y) Then
    '        For X = 1 To 4  'Earnings
    '            errMsg = False
    '            If Len("medOthEarn" & X & "WK" & Y) = 0 Then
    '                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '                MsgBox "'Any Other Earnings' for column " & X & " for Week " & Y & " cannot be blank.", vbExclamation
    '                errMsg = True
    '            ElseIf Not IsNumeric("medOthEarn" & X & "WK" & Y) = 0 Then
    '                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '                MsgBox "Invalid 'Any Other Earnings' for column " & X & " for Week " & Y & ".", vbExclamation
    '                errMsg = True
    '            ElseIf Val("medOthEarn" & X & "WK" & Y) < 0 Then
    '                tbF7Sections.SelectedItem = tbF7Sections.Tabs(3)
    '                MsgBox "'Any Other Earnings' for column " & X & " for Week " & Y & " cannot be less than 0.", vbExclamation
    '                errMsg = True
    '            End If
    '            If errMsg Then
    '                Select Case Y
    '                    Case 1
    '                        Select Case X
    '                            Case 1: medOthEarn1WK1.SetFocus
    '                            Case 2: medOthEarn2WK1.SetFocus
    '                            Case 3: medOthEarn3WK1.SetFocus
    '                            Case 4: medOthEarn4WK1.SetFocus
    '                        End Select
    '                    Case 2
    '                        Select Case X
    '                            Case 1: medOthEarn1WK2.SetFocus
    '                            Case 2: medOthEarn2WK2.SetFocus
    '                            Case 3: medOthEarn3WK2.SetFocus
    '                            Case 4: medOthEarn4WK2.SetFocus
    '                        End Select
    '                    Case 3
    '                        Select Case X
    '                            Case 1: medOthEarn1WK3.SetFocus
    '                            Case 2: medOthEarn2WK3.SetFocus
    '                            Case 3: medOthEarn3WK3.SetFocus
    '                            Case 4: medOthEarn4WK3.SetFocus
    '                        End Select
    '                    Case 4
    '                        Select Case X
    '                            Case 1: medOthEarn1WK4.SetFocus
    '                            Case 2: medOthEarn2WK4.SetFocus
    '                            Case 3: medOthEarn3WK4.SetFocus
    '                            Case 4: medOthEarn4WK4.SetFocus
    '                        End Select
    '                End Select
    '                Exit Function
    '            End If
    '        Next
    '    End If
    'Next
    
    'Section I: Work Schedule
    If optWorkSchedule(0).Value = False And optWorkSchedule(1).Value = False And optWorkSchedule(2).Value = False Then
        tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
        MsgBox "One of the 'Work Schedule (Regular/Repeating Rotational Shift Worker/Varied or Irregular)' indicator selection is required.", vbExclamation
        'optWorkSchedule(0).SetFocus
        'Exit Function
    End If
    
    If optWorkSchedule(0).Value = True Then
        For X = 0 To 6
            If Len(Trim(medHours(X).Text)) > 0 Then
                If Not IsNumeric(medHours(X).Text) Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "Invalid " & lblWeekDay(X).Caption & " Hours.", vbExclamation
                    medHours(X).SetFocus
                    Exit Function
                ElseIf Val(medHours(X).Text) <= 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox lblWeekDay(X).Caption & " Hours cannot be less or equal to 0.", vbExclamation
                    medHours(X).SetFocus
                    Exit Function
                End If
            End If
        Next
        
        flgHoursEntered = False
        For X = 0 To 6
            If Len(Trim(medHours(X).Text)) > 0 Then
                flgHoursEntered = True
            End If
        Next
        If flgHoursEntered = False Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "No Regular Schedule Hours entered. At least one day is required to be filled in.", vbExclamation
            'medHours(0).SetFocus
            'Exit Function
        End If
    End If
    
    If optWorkSchedule(1).Value = True Then
        If Len(Trim(medNoDayOn.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Number of Days On' is required for 'Repeating Rotational Shift Worker'.", vbExclamation
            'medNoDayOn.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medNoDayOn.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Number of Days On'.", vbExclamation
            medNoDayOn.SetFocus
            Exit Function
        ElseIf Val(medNoDayOn.Text) <= 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Number of Days On' cannot be less or equal to 0.", vbExclamation
            medNoDayOn.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medNoDayOff.Text)) = 0 Then
            MsgBox "'Number of Days Off' is required for 'Repeating Rotational Shift Worker'.", vbExclamation
            'medNoDayOff.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medNoDayOff.Text) Then
            MsgBox "Invalid 'Number of Days Off'.", vbExclamation
            medNoDayOff.SetFocus
            Exit Function
        ElseIf Val(medNoDayOff.Text) <= 0 Then
            MsgBox "'Number of Days Off' cannot be less or equal to 0.", vbExclamation
            medNoDayOff.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medHrsShift.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Hours per Shift(s)' is required for 'Repeating Rotational Shift Worker'.", vbExclamation
            'medHrsShift.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medHrsShift.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Hours per Shift(s)'.", vbExclamation
            medHrsShift.SetFocus
            Exit Function
        ElseIf Val(medHrsShift.Text) <= 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Hours per Shift(s)' cannot be less or equal to 0.", vbExclamation
            medHrsShift.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medNoWksCycle.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Number of Weeks in Cycle' is required for 'Repeating Rotational Shift Worker'.", vbExclamation
            'medNoWksCycle.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medNoWksCycle.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Number of Weeks in Cycle'.", vbExclamation
            medNoWksCycle.SetFocus
            Exit Function
        ElseIf Val(medNoWksCycle.Text) <= 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Number of Weeks in Cycle' cannot be less or equal to 0.", vbExclamation
            medNoWksCycle.SetFocus
            Exit Function
        End If
    End If
    
    If optWorkSchedule(2).Value = True Then
        'Week 1
        If Len(Trim(dlpWk1FDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "From Date for Week 1 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk1FDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk1FDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid From Date for Week 1 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk1FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(dlpWk1TDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "To Date for Week 1 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk1TDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk1TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid To Date for Week 1 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk1TDate.SetFocus
            Exit Function
        End If
        
        If CVDate(dlpWk1FDate.Text) > CVDate(dlpWk1TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'From Date for Week 1' cannot be greater than 'To Date for Week 1' for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk1FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medTotHrsWrkWK1.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Hours Worked' for Week 1 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotHrsWrkWK1.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotHrsWrkWK1.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Hours Worked' for Week 1 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotHrsWrkWK1.SetFocus
            Exit Function
        ElseIf Val(medTotHrsWrkWK1.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotHrsWrkWK1.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Hours Worked' for Week 1 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotHrsWrkWK1.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Hours Worked' for Week 1 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotHrsWrkWK1.SetFocus
                Exit Function
            End If
        End If
    
        If Len(Trim(medTotShiftsWrkWK1.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Shifts Worked' for Week 1 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotShiftsWrkWK1.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotShiftsWrkWK1.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Shifts Worked' for Week 1 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotShiftsWrkWK1.SetFocus
            Exit Function
        ElseIf Val(medTotShiftsWrkWK1.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotShiftsWrkWK1.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Shifts Worked' for Week 1 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotShiftsWrkWK1.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Shifts Worked' for Week 1 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotShiftsWrkWK1.SetFocus
                Exit Function
            End If
        End If
        
        'Week 2
        If Len(Trim(dlpWk2FDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "From Date for Week 2 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk2FDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk2FDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid From Date for Week 2 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk2FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(dlpWk2TDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "To Date for Week 2 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk2TDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk2TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid To Date for Week 2 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk2TDate.SetFocus
            Exit Function
        End If
        
        If CVDate(dlpWk2FDate.Text) > CVDate(dlpWk2TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'From Date for Week 2' cannot be greater than 'To Date for Week 2' for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk2FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medTotHrsWrkWK2.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Hours Worked' for Week 2 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotHrsWrkWK2.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotHrsWrkWK2.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Hours Worked' for Week 2 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotHrsWrkWK2.SetFocus
            Exit Function
        ElseIf Val(medTotHrsWrkWK2.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotHrsWrkWK2.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Hours Worked' for Week 2 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotHrsWrkWK2.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Hours Worked' for Week 2 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotHrsWrkWK2.SetFocus
                Exit Function
            End If
        End If
    
        If Len(Trim(medTotShiftsWrkWK2.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Shifts Worked' for Week 2 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotShiftsWrkWK2.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotShiftsWrkWK2.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Shifts Worked' for Week 2 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotShiftsWrkWK2.SetFocus
            Exit Function
        ElseIf Val(medTotShiftsWrkWK2.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotShiftsWrkWK2.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Shifts Worked' for Week 2 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotShiftsWrkWK2.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Shifts Worked' for Week 2 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotShiftsWrkWK2.SetFocus
                Exit Function
            End If
        End If
        
        'Week 3
        If Len(Trim(dlpWk3FDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "From Date for Week 3 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk3FDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk3FDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid From Date for Week 3 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk3FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(dlpWk3TDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "To Date for Week 3 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk3TDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk3TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid To Date for Week 3 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk3TDate.SetFocus
            Exit Function
        End If
        
        If CVDate(dlpWk3FDate.Text) > CVDate(dlpWk3TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'From Date for Week 3' cannot be greater than 'To Date for Week 3' for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk3FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medTotHrsWrkWK3.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Hours Worked' for Week 3 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotHrsWrkWK3.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotHrsWrkWK3.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Hours Worked' for Week 3 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotHrsWrkWK3.SetFocus
            Exit Function
        ElseIf Val(medTotHrsWrkWK3.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotHrsWrkWK3.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Hours Worked' for Week 3 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotHrsWrkWK3.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Hours Worked' for Week 3 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotHrsWrkWK3.SetFocus
                Exit Function
            End If
        End If
    
        If Len(Trim(medTotShiftsWrkWK3.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Shifts Worked' for Week 3 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotShiftsWrkWK3.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotShiftsWrkWK3.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Shifts Worked' for Week 3 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotShiftsWrkWK3.SetFocus
            Exit Function
        ElseIf Val(medTotShiftsWrkWK3.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotShiftsWrkWK3.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Shifts Worked' for Week 3 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotShiftsWrkWK3.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Shifts Worked' for Week 3 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotShiftsWrkWK3.SetFocus
                Exit Function
            End If
        End If
        
        'Week 4
        If Len(Trim(dlpWk4FDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "From Date for Week 4 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk4FDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk4FDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid From Date for Week 4 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk4FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(dlpWk4TDate.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "To Date for Week 4 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'dlpWk4TDate.SetFocus
            'Exit Function
        ElseIf Not IsDate(dlpWk4TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid To Date for Week 4 for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk4TDate.SetFocus
            Exit Function
        End If
        
        If CVDate(dlpWk4FDate.Text) > CVDate(dlpWk4TDate.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'From Date for Week 4' cannot be greater than 'To Date for Week 4' for 'Varied or Irregular Work Schedule'.", vbExclamation
            dlpWk4FDate.SetFocus
            Exit Function
        End If
    
        If Len(Trim(medTotHrsWrkWK4.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Hours Worked' for Week 4 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotHrsWrkWK4.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotHrsWrkWK4.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Hours Worked' for Week 4 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotHrsWrkWK4.SetFocus
            Exit Function
        ElseIf Val(medTotHrsWrkWK4.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotHrsWrkWK4.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Hours Worked' for Week 4 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotHrsWrkWK4.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Hours Worked' for Week 4 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotHrsWrkWK4.SetFocus
                Exit Function
            End If
        End If
    
        If Len(Trim(medTotShiftsWrkWK4.Text)) = 0 Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "'Total Shifts Worked' for Week 4 is required for 'Varied or Irregular Work Schedule'.", vbExclamation
            'medTotShiftsWrkWK4.SetFocus
            'Exit Function
        ElseIf Not IsNumeric(medTotShiftsWrkWK4.Text) Then
            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
            MsgBox "Invalid 'Total Shifts Worked' for Week 4 for 'Varied or Irregular Work Schedule'.", vbExclamation
            medTotShiftsWrkWK4.SetFocus
            Exit Function
        ElseIf Val(medTotShiftsWrkWK4.Text) <= 0 Then
            If glbCompSerial = "S/N - 2295W" Then 'City of St. Thomas Ticket #28985 Franks 07/26/2016 - allow 0
                If Val(medTotShiftsWrkWK4.Text) < 0 Then
                    tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                    MsgBox "'Total Shifts Worked' for Week 4 for 'Varied or Irregular Work Schedule' cannot be less than 0.", vbExclamation
                    medTotShiftsWrkWK4.SetFocus
                    Exit Function
                End If
            Else
                tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
                MsgBox "'Total Shifts Worked' for Week 4 for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
                medTotShiftsWrkWK4.SetFocus
                Exit Function
            End If
        End If
        
    End If
    
    'If optWorkSchedule(2).Value = True Then
    '    errMsg = False
    '    For X = 1 To 4
    '        If Len(Trim("dlpWk" & X & "Fdate")) = 0 Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "From Date for Week " & X & " is required for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        ElseIf Not IsDate("dlpWk" & X & "Fdate") Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "Invalid From Date for Week " & X & " for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        End If
    '        If errMsg Then
    '            Select Case X
    '                Case 1: dlpWk1FDate.SetFocus
    '                Case 2: dlpWk2FDate.SetFocus
    '                Case 3: dlpWk3FDate.SetFocus
    '                Case 4: dlpWk4FDate.SetFocus
    '            End Select
    '            Exit Function
    '        End If
    '
    '        errMsg = False
    '        If Len(Trim("dlpWk" & X & "Tdate")) = 0 Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "To Date for Week " & X & " is required for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        ElseIf Not IsDate("dlpWk" & X & "Tdate") Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "Invalid To Date for Week " & X & " for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        End If
    '        If errMsg Then
    '            Select Case X
    '                Case 1: dlpWk1TDate.SetFocus
    '                Case 2: dlpWk2TDate.SetFocus
    '                Case 3: dlpWk3TDate.SetFocus
    '                Case 4: dlpWk4TDate.SetFocus
    '            End Select
    '            Exit Function
    '        End If
    '
    '        errMsg = False
    '        If CVDate("dlpWk" & X & "Fdate") > CVDate("dlpWk" & X & "Tdate") Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "'From Date for Week " & X & "' cannot be greater than 'To Date for Week " & X & "' for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        End If
    '        If errMsg Then
    '            Select Case X
    '                Case 1: dlpWk1FDate.SetFocus
    '                Case 2: dlpWk2FDate.SetFocus
    '                Case 3: dlpWk3FDate.SetFocus
    '                Case 4: dlpWk4FDate.SetFocus
    '            End Select
    '            Exit Function
    '        End If
    '
    '        errMsg = False
    '        If Len(Trim("medTotHrsWrkWK" & X)) = 0 Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "'Total Hours Worked' for Week " & X & " is required for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        ElseIf Not IsNumeric("medTotHrsWrkWK" & X) Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "Invalid 'Total Hours Worked' for Week " & X & " for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        ElseIf Val("medTotHrsWrkWK" & X) <= 0 Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "'Total Hours Worked' for Week " & X & " for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
    '            errMsg = True
    '        End If
    '        If errMsg Then
    '            Select Case X
    '                Case 1: medTotHrsWrkWK1.SetFocus
    '                Case 2: medTotHrsWrkWK2.SetFocus
    '                Case 3: medTotHrsWrkWK3.SetFocus
    '                Case 4: medTotHrsWrkWK4.SetFocus
    '            End Select
    '            Exit Function
    '        End If
    '
    '        errMsg = False
    '        If Len(Trim("medTotShiftsWrkWK" & X)) = 0 Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "'Total Shifts Worked' for Week " & X & " is required for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        ElseIf Not IsNumeric("medTotShiftsWrkWK" & X) Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "Invalid 'Total Shifts Worked' for Week " & X & " for 'Varied or Irregular Work Schedule'.", vbExclamation
    '            errMsg = True
    '        ElseIf Val("medTotShiftsWrkWK" & X) <= 0 Then
    '            tbF7Sections.SelectedItem = tbF7Sections.Tabs(4)
    '            MsgBox "'Total Shifts Worked' for Week " & X & " for 'Varied or Irregular Work Schedule' cannot be less or equal to 0.", vbExclamation
    '            errMsg = True
    '        End If
    '        If errMsg Then
    '            Select Case X
    '                Case 1: medTotShiftsWrkWK1.SetFocus
    '                Case 2: medTotShiftsWrkWK2.SetFocus
    '                Case 3: medTotShiftsWrkWK3.SetFocus
    '                Case 4: medTotShiftsWrkWK4.SetFocus
    '            End Select
    '            Exit Function
    '        End If
    '    Next
    'End If
End If

chkHSInjuryF7Sections = True

Exit Function

chkHSInjuryF7_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkInjuryF7", "HR_OHS_FORM7_SECTIONS", "edit/Add")

If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub chkWrittenOfferAttached_Click()
    If chkWrittenOfferAttached.Value = 1 Then
        cmdImport.Enabled = True
    Else
        cmdImport.Enabled = False
    End If
End Sub

Private Sub cmbFilledByName_Click()
    If cmbFilledByName.Text = "" Then
        txtFilledByName.Text = ""
        txtFilledByTitle.Text = ""
        medFilledByTelephone.Text = ""
        medFilledByTelephone.Mask = "(###) ###-####  Ext(######)"
    End If
    
    If cmbFilledByName.ListIndex <> -1 Then
        txtFilledByName.Text = cmbFilledByName.Text
        
        'Retrieve selected Name's Title and Phone #
        Call Get_FilledBy_Details
    Else
        txtFilledByName.Text = ""
        txtFilledByTitle.Text = ""
        medFilledByTelephone.Text = ""
        medFilledByTelephone.Mask = "(###) ###-####  Ext(######)"
    End If
End Sub

Private Sub cmbFilledByName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbFilledByName_LostFocus()
    If cmbFilledByName.Text = "" Then
        txtFilledByName.Text = ""
        txtFilledByTitle.Text = ""
        medFilledByTelephone.Text = ""
        medFilledByTelephone.Mask = "(###) ###-####  Ext(######)"
    End If
End Sub

Public Sub cmdClose_Click()

    On Error GoTo err_Unload
    
    'glbTERM_ID = 0
    'glbTran_ID = 0
    'glbTran_Seq = 0
    glbOnTop = ""
    
    
    Unload Me
    
    Exit Sub
    
err_Unload:
    Unload Me
    Resume Next
    Unload Me

End Sub

'Form 7 - Additional Sections - Section F
Private Sub cmdImport_Click()
Dim xID
    glbDocNewRecord = False
    glbDocName = "INJURYWF7_WRITTENOFR"
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        glbDocKey = 0
        glbJob = ""
        glbDocTmp = ""
    Else
        glbDocKey = rsDATA("F7_ID")
        glbJob = rsDATA("F7_CASE")
        'glbDocTmp = rsDATA1("EC_DOCKEY")
    End If

    frmInAttachment.Show 1
    DoEvents
    
    Call DispimgIcon(Me, "frmEInjF7Sections")
    
    If chkWrittenOfferAttached.Value = 1 Then
        cmdImport.Enabled = True
    Else
        cmdImport.Enabled = False
    End If
End Sub

Private Function FldList()
Dim SQLQ

SQLQ = ""
SQLQ = SQLQ & "* "

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList = SQLQ

End Function

Private Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_OHS_FORM7_SECTIONS "
    SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " AND F7_CASE = " & frmEHSINJURYWF7.lblIncidentNo.Caption
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OHS_FORM7_SECTIONS "
    SQLQ = SQLQ & "WHERE F7_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND F7_CASE = " & frmEHSINJURYWF7.lblIncidentNo.Caption
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
'SQLQ = SQLQ & " ORDER BY F7_CASE DESC"
Data1.RecordSource = SQLQ
Data1.Refresh

Call Display_Value

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Injury Form 7 Sections", "EERetrieve", "SELECT")

Resume Next

Exit Function
'
End Function

Sub Display_Value()
    Dim SQLQ
    
    'Form 7 - Additional Sections - Section F
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        'Me.cmdModify_Click
        Exit Sub
    End If
        
'    If glbtermopen Then
'        SQLQ = "SELECT " & FldList & " FROM Term_OHS_FORM7_SECTIONS "
'        SQLQ = SQLQ & "WHERE F7_CASE=" & Data1.Recordset!F7_CASE
'        SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
'        SQLQ = SQLQ & " AND TERM_SEQ=" & glbTERM_Seq
'        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'    Else
'        SQLQ = "SELECT " & FldList & " FROM HR_OHS_FORM7_SECTIONS "
'        SQLQ = SQLQ & "WHERE F7_CASE = " & Data1.Recordset!F7_CASE
'        SQLQ = SQLQ & " AND F7_EMPNBR =" & glbLEE_ID
'
'        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
'        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    End If
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
        
    Call SET_UP_MODE
    
    'Me.cmdModify_Click

End Sub

Private Sub cmdRefresh_Click()
    Dim Response%
    Dim flgSuccess As Boolean
    
    Response% = MsgBox("'H. Additional Wage Information' and 'I. Work Schedule' will be refreshed with employee's current data. Are you sure you want to continue with the refresh?", vbQuestion + vbYesNo, "info:HR - Refresh Form 7 Data")
    If Response% = IDYES Then    ' Evaluate response
        Screen.MousePointer = HOURGLASS
        
        'Call procedure to refresh the Tax exception amounts and dates
        flgSuccess = Refresh_Form7_Data
        
        Screen.MousePointer = DEFAULT
        
        If flgSuccess Then
            MsgBox "'H. Additional Wage Information' and 'I. Work Schedule' data refreshed successfully.", vbInformation, "info:HR - Refresh Form 7 Data"
        Else
            MsgBox "'H. Additional Wage Information' and 'I. Work Schedule' data was NOT refreshed.", vbExclamation, "info:HR - Refresh Form 7 Data"
        End If
    End If

End Sub

Private Sub comOtherEarnings1_Change()
    txtOtherEarnings1.Text = comOtherEarnings1.Text
End Sub

Private Sub comOtherEarnings1_Click()
    txtOtherEarnings1.Text = comOtherEarnings1.Text
End Sub

Private Sub comOtherEarnings1_LostFocus()
    txtOtherEarnings1.Text = comOtherEarnings1.Text
End Sub

Private Sub comOtherEarnings2_Change()
    txtOtherEarnings2.Text = comOtherEarnings2.Text
End Sub

Private Sub comOtherEarnings2_Click()
    txtOtherEarnings2.Text = comOtherEarnings2.Text
End Sub

Private Sub comOtherEarnings2_LostFocus()
    txtOtherEarnings2.Text = comOtherEarnings2.Text
End Sub

Private Sub comOtherEarnings3_Change()
    txtOtherEarnings3.Text = comOtherEarnings3.Text
End Sub

Private Sub comOtherEarnings3_Click()
    txtOtherEarnings3.Text = comOtherEarnings3.Text
End Sub

Private Sub comOtherEarnings3_LostFocus()
    txtOtherEarnings3.Text = comOtherEarnings3.Text
End Sub

Private Sub comOtherEarnings4_Change()
    txtOtherEarnings4.Text = comOtherEarnings4.Text
End Sub

Private Sub comOtherEarnings4_Click()
    txtOtherEarnings4.Text = comOtherEarnings4.Text
End Sub

Private Sub comOtherEarnings4_LostFocus()
    txtOtherEarnings4.Text = comOtherEarnings4.Text
End Sub

Private Sub dlpOtherEarnFromWK1_LostFocus()
    If IsDate(dlpOtherEarnFromWK1.Text) And dlpOtherEarnToWK1.Text = "" And _
        dlpOtherEarnFromWK2.Text = "" And dlpOtherEarnToWK2.Text = "" And _
        dlpOtherEarnFromWK3.Text = "" And dlpOtherEarnToWK3.Text = "" And _
        dlpOtherEarnFromWK4.Text = "" And dlpOtherEarnToWK4.Text = "" Then
        
        'Compute dates for the rest of the weeks
        'Week 1 - Compute To Date using From Date
        dlpOtherEarnToWK1.Text = DateAdd("d", 6, CVDate(dlpOtherEarnFromWK1.Text))
        
        'Week 2 - Compute To using the previous From Date and then From Date
        dlpOtherEarnToWK2.Text = DateAdd("d", -1, CVDate(dlpOtherEarnFromWK1.Text))
        dlpOtherEarnFromWK2.Text = DateAdd("d", -6, CVDate(dlpOtherEarnToWK2.Text))
        
        'Week 3 - Compute To using the previous From and then From Date
        dlpOtherEarnToWK3.Text = DateAdd("d", -1, CVDate(dlpOtherEarnFromWK2.Text))
        dlpOtherEarnFromWK3.Text = DateAdd("d", -6, CVDate(dlpOtherEarnToWK3.Text))
        
        'Week 4 - Compute To using the previous From and then From Date
        dlpOtherEarnToWK4.Text = DateAdd("d", -1, CVDate(dlpOtherEarnFromWK3.Text))
        dlpOtherEarnFromWK4.Text = DateAdd("d", -6, CVDate(dlpOtherEarnToWK4.Text))
        
        'Section 1: Option C - Re-populate if selected
        If optWorkSchedule(2).Value = True Then
            If IsDate(dlpOtherEarnFromWK1.Text) Then dlpWk1FDate.Text = dlpOtherEarnFromWK1.Text
            If IsDate(dlpOtherEarnToWK1.Text) Then dlpWk1TDate.Text = dlpOtherEarnToWK1.Text
            If IsDate(dlpOtherEarnFromWK2.Text) Then dlpWk2FDate.Text = dlpOtherEarnFromWK2.Text
            If IsDate(dlpOtherEarnToWK2.Text) Then dlpWk2TDate.Text = dlpOtherEarnToWK2.Text
            If IsDate(dlpOtherEarnFromWK3.Text) Then dlpWk3FDate.Text = dlpOtherEarnFromWK3.Text
            If IsDate(dlpOtherEarnToWK3.Text) Then dlpWk3TDate.Text = dlpOtherEarnToWK3.Text
            If IsDate(dlpOtherEarnFromWK4.Text) Then dlpWk4FDate.Text = dlpOtherEarnFromWK4.Text
            If IsDate(dlpOtherEarnToWK4.Text) Then dlpWk4TDate.Text = dlpOtherEarnToWK4.Text
        End If
    End If
End Sub

Private Sub dlpRetLostDate_Change(Index As Integer)
    If IsDate(dlpRetLostDate(1).Text) Then
        optRegMod(0).Enabled = True
        optRegMod(1).Enabled = True
    Else
        optRegMod(0).Enabled = False
        optRegMod(1).Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    glbOnTop = "frmEInjF7Sections"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmEInjF7Sections"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim rsTERM As New ADODB.Recordset
Dim X%, SQLQ

glbOnTop = "frmEInjF7Sections"

'Call setCaption(lblTitle(1))

Screen.MousePointer = HOURGLASS

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If
Screen.MousePointer = DEFAULT

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblTitle(0).Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(glbLEE_ID)
End If

'Show Section E first time
frFReturnToWork.Visible = False
frHAdditionalWage.Visible = False
frIWorkSchedule.Visible = False
frKAdditionalInfo.Visible = False
frJFilledBy.Visible = False
frELostTime.Visible = True
frELostTime.Top = 1880
frELostTime.Left = 360

'Ticket #22682 - Release 8.0 - Load Filled By Names
Call Populate_FilledBy_Names

If EERetrieve() = False Then Exit Sub

'Call Display_Value

MDIMain.panHelp(1).Caption = " "

Call INI_Controls(Me)

comOtherEarnings1.Clear
comOtherEarnings2.Clear
comOtherEarnings3.Clear
comOtherEarnings4.Clear

comOtherEarnings1 = Trim(txtOtherEarnings1.Text)
comOtherEarnings2 = Trim(txtOtherEarnings2.Text)
comOtherEarnings3 = Trim(txtOtherEarnings3.Text)
comOtherEarnings4 = Trim(txtOtherEarnings4.Text)

'Populate Other Earnings
Call Populate_OtherEarnings_ComboBox


End Sub

Private Sub Form_Unload(Cancel As Integer)
    glbOnTop = ""
End Sub

Private Sub cmdOK_Click()
Dim X
Dim xBPart As String
Dim otherEarnlist

On Error GoTo Add_Err

If Not chkHSInjuryF7Sections() Then Exit Sub

rsDATA.Requery

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh


'Repopoulate Other Earning combobox incase new value entered.
'Clear the existing list
comOtherEarnings1.Clear
comOtherEarnings2.Clear
comOtherEarnings3.Clear
comOtherEarnings4.Clear

'Reset the data value into the combobox control
comOtherEarnings1 = Trim(txtOtherEarnings1.Text)
comOtherEarnings2 = Trim(txtOtherEarnings2.Text)
comOtherEarnings3 = Trim(txtOtherEarnings3.Text)
comOtherEarnings4 = Trim(txtOtherEarnings4.Text)

'Repopulate the list using the MTF file including the new value entered by the user.
Call Populate_OtherEarnings_ComboBox
'otherEarnlist = OtherEarningList



Call SET_UP_MODE

MsgBox "Changes saved successfully.", vbInformation, "info:HR - Form 7 - Additional Sections"

Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_FORM7_SECTIONS", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
 ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateTransEmp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_HSW7Injury And glbWSIBModule
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = xUpdateable
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

UpdateState = OPENING
TF = True

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

Call ST_UPD_MODE(TF)

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


frFReturnToWork.Enabled = TF
frHAdditionalWage.Enabled = TF
frIWorkSchedule.Enabled = TF
frKAdditionalInfo.Enabled = TF
frJFilledBy.Enabled = TF
frELostTime.Enabled = TF
cmdOK.Enabled = TF
cmdRefresh.Enabled = TF

If Data1.Recordset.BOF And Data1.Recordset.EOF Then 'Add by Frank 8/21/2001
    'cmdModify.Enabled = False
Else
    'Me.cmdModify_Click
End If

'Return to Work declined
glbJob = ""
glbSDate = "01/01/1900"
glbDocKey = 0
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    If Data1.Recordset("F7_DECLINE_ATTACHED") = True Or chkWrittenOfferAttached.Value = 1 Then
        glbJob = Data1.Recordset("F7_CASE")
        glbSDate = Data1.Recordset("F7_OCCDATE")
        glbDocKey = IIf(IsNull(Data1.Recordset("F7_DOCKEY")), "", Data1.Recordset("F7_DOCKEY"))
        'glbDocTmp = IIf(IsNull(Data1.Recordset("F7_DOCKEY")), "", Data1.Recordset("F7_DOCKEY"))
    Else
        glbJob = ""
        glbSDate = ""
        glbDocKey = ""
        glbDocTmp = ""
    End If
End If

glbDocName = "INJURYWF7_WRITTENOFR"
If gsAttachment_DB Then
    Call DispimgIcon(Me, "frmEInjF7Sections")
    If gSec_Upd_HSW7Injury And glbWSIBModule And Not glbtermopen Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
            If Data1.Recordset("F7_DECLINE_ATTACHED") = True Then
                cmdImport.Enabled = True
            Else
                cmdImport.Enabled = False
            End If
        End If
    End If
End If
 
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEInjF7Sections")
    Call FillMemoFile(SQLQ, "INJURYWF7_WRITTENOFR")
End Sub


Private Sub medHourLastWorked_GotFocus()
    medHourLastWorked.Mask = "##:##"
End Sub

Private Sub medHourLastWorked_LostFocus()
    medHourLastWorked.Mask = ""
End Sub

Private Sub medNormTimeFrom_GotFocus()
    medNormTimeFrom.Mask = "##:##"
End Sub

Private Sub medNormTimeFrom_LostFocus()
    medNormTimeFrom.Mask = ""
End Sub


Private Sub medNormTimeTo_GotFocus()
    medNormTimeTo.Mask = "##:##"
End Sub

Private Sub medNormTimeTo_LostFocus()
    medNormTimeTo.Mask = ""
End Sub

Private Sub optAccDecl_Click(Index As Integer)
    If optAccDecl(0).Value = True Then
        txtAccDecl.Text = "A"
    ElseIf optAccDecl(1).Value = True Then
        txtAccDecl.Text = "D"
    Else
        txtAccDecl.Text = ""
    End If
End Sub

Private Sub optBeingPaidYN_Click(Index As Integer)
    If optBeingPaidYN(0).Value = True Then
        txtBeingPaidYN.Text = "1"
    ElseIf optBeingPaidYN(1).Value Then
        txtBeingPaidYN.Text = "0"
    Else
        txtBeingPaidYN.Text = ""
    End If
End Sub

Private Sub optConfirmedBy_Click(Index As Integer)
    If optConfirmedBy(0).Value = True Then
        txtConfirmedBy.Text = "M"
    ElseIf optConfirmedBy(1).Value = True Then
        txtConfirmedBy.Text = "O"
    Else
        txtConfirmedBy.Text = ""
    End If
End Sub

Private Sub optDiscussed_Click(Index As Integer)
    If optDiscussed(0).Value = True Then
        txtDiscussed.Text = "1"
    ElseIf optDiscussed(1).Value Then
        txtDiscussed.Text = "0"
    Else
        txtDiscussed.Text = ""
    End If
End Sub

Private Sub optFullRegOther_Click(Index As Integer)
    If optFullRegOther(0).Value = True Then
        txtFullRegOther.Text = "F"
    ElseIf optFullRegOther(1).Value = True Then
        txtFullRegOther.Text = "O"
    Else
        txtFullRegOther.Text = ""
    End If
End Sub

Private Sub optLimitations_Click(Index As Integer)
    If optLimitations(0).Value = True Then
        txtLimitations.Text = "1"
    ElseIf optLimitations(1).Value Then
        txtLimitations.Text = "0"
    Else
        txtLimitations.Text = ""
    End If
End Sub

Private Sub optLostTime_Click(Index As Integer)
    If optLostTime(0).Value = True Then
        txtLostTime.Text = "R"
    ElseIf optLostTime(1).Value = True Then
        txtLostTime.Text = "M"
    ElseIf optLostTime(2).Value = True Then
        txtLostTime.Text = "L"
    Else
        txtLostTime.Text = ""
    End If
End Sub

Private Sub optNormFTimeLastWorkAP_Click(Index As Integer)
    If optNormFTimeLastWorkAP(0).Value = True Then
        txtNormFTimeLastWorkAP.Text = "A"
    ElseIf optNormFTimeLastWorkAP(1).Value = True Then
        txtNormFTimeLastWorkAP.Text = "P"
    Else
        txtNormFTimeLastWorkAP.Text = ""
    End If
End Sub

Private Sub optNormTTimeLastWorkAP_Click(Index As Integer)
    If optNormTTimeLastWorkAP(0).Value = True Then
        txtNormTTimeLastWorkAP.Text = "A"
    ElseIf optNormTTimeLastWorkAP(1).Value = True Then
        txtNormTTimeLastWorkAP.Text = "P"
    Else
        txtNormTTimeLastWorkAP.Text = ""
    End If
End Sub

Private Sub optOffered_Click(Index As Integer)
    If optOffered(0).Value = True Then
        txtOffered.Text = "1"
    ElseIf optOffered(1).Value Then
        txtOffered.Text = "0"
    Else
        txtOffered.Text = ""
    End If
End Sub

Private Sub optRegMod_Click(Index As Integer)
    If optRegMod(0).Value = True Then
        txtRegMod.Text = "R"
    ElseIf optRegMod(1).Value = True Then
        txtRegMod.Text = "M"
    Else
        txtRegMod.Text = ""
    End If
End Sub

Private Sub optResponsible_Click(Index As Integer)
    If optResponsible(0).Value = True Then
        txtResponsible.Text = "M"
    ElseIf optResponsible(1).Value = True Then
        txtResponsible.Text = "O"
    Else
        txtResponsible.Text = ""
    End If
End Sub

Private Sub optTimeLastWorkAP_Click(Index As Integer)
    If optTimeLastWorkAP(0).Value = True Then
        txtTimeLastWorkAP.Text = "A"
    ElseIf optTimeLastWorkAP(1).Value = True Then
        txtTimeLastWorkAP.Text = "P"
    Else
        txtTimeLastWorkAP.Text = ""
    End If
End Sub

Private Sub optVacPerctYN_Click(Index As Integer)
    If optVacPerctYN(0).Value = True Then
        txtVacPerctYN.Text = "1"
    ElseIf optVacPerctYN(1).Value Then
        txtVacPerctYN.Text = "0"
    Else
        txtVacPerctYN.Text = ""
    End If
End Sub

Private Sub optWorkSchedule_Click(Index As Integer)
    Dim X As Integer
    
    If optWorkSchedule(0).Value = True Then
        txtWorkSchedule.Text = "N"
        
        'Clear the contents of Option B and C.
        'Option B.
        medNoDayOn.Text = ""
        medNoDayOff.Text = ""
        medHrsShift.Text = ""
        medNoWksCycle.Text = ""
        
        'Option C
        dlpWk1FDate.Text = ""
        dlpWk1TDate.Text = ""
        dlpWk2FDate.Text = ""
        dlpWk2TDate.Text = ""
        dlpWk3FDate.Text = ""
        dlpWk3TDate.Text = ""
        dlpWk4FDate.Text = ""
        dlpWk4TDate.Text = ""
        
        medTotHrsWrkWK1.Text = ""
        medTotHrsWrkWK2.Text = ""
        medTotHrsWrkWK3.Text = ""
        medTotHrsWrkWK4.Text = ""
        medTotShiftsWrkWK1.Text = ""
        medTotShiftsWrkWK2.Text = ""
        medTotShiftsWrkWK3.Text = ""
        medTotShiftsWrkWK4.Text = ""
    ElseIf optWorkSchedule(1).Value = True Then
        txtWorkSchedule.Text = "R"
    
        'Clear the contents of Option B and C.
        'Option A
        For X = 0 To 6
            medHours(X).Text = ""
        Next
        
        'Option C
        dlpWk1FDate.Text = ""
        dlpWk1TDate.Text = ""
        dlpWk2FDate.Text = ""
        dlpWk2TDate.Text = ""
        dlpWk3FDate.Text = ""
        dlpWk3TDate.Text = ""
        dlpWk4FDate.Text = ""
        dlpWk4TDate.Text = ""
        
        medTotHrsWrkWK1.Text = ""
        medTotHrsWrkWK2.Text = ""
        medTotHrsWrkWK3.Text = ""
        medTotHrsWrkWK4.Text = ""
        medTotShiftsWrkWK1.Text = ""
        medTotShiftsWrkWK2.Text = ""
        medTotShiftsWrkWK3.Text = ""
        medTotShiftsWrkWK4.Text = ""
    ElseIf optWorkSchedule(2).Value = True Then
        txtWorkSchedule.Text = "I"
    
        'Clear the contents of Option A and B.
        'Option A
        For X = 0 To 6
            medHours(X).Text = ""
        Next
        
        'Option B.
        medNoDayOn.Text = ""
        medNoDayOff.Text = ""
        medHrsShift.Text = ""
        medNoWksCycle.Text = ""
        
        'Option C - Re-populate
        If IsDate(dlpOtherEarnFromWK1.Text) Then dlpWk1FDate.Text = dlpOtherEarnFromWK1.Text
        If IsDate(dlpOtherEarnToWK1.Text) Then dlpWk1TDate.Text = dlpOtherEarnToWK1.Text
        If IsDate(dlpOtherEarnFromWK2.Text) Then dlpWk2FDate.Text = dlpOtherEarnFromWK2.Text
        If IsDate(dlpOtherEarnToWK2.Text) Then dlpWk2TDate.Text = dlpOtherEarnToWK2.Text
        If IsDate(dlpOtherEarnFromWK3.Text) Then dlpWk3FDate.Text = dlpOtherEarnFromWK3.Text
        If IsDate(dlpOtherEarnToWK3.Text) Then dlpWk3TDate.Text = dlpOtherEarnToWK3.Text
        If IsDate(dlpOtherEarnFromWK4.Text) Then dlpWk4FDate.Text = dlpOtherEarnFromWK4.Text
        If IsDate(dlpOtherEarnToWK4.Text) Then dlpWk4TDate.Text = dlpOtherEarnToWK4.Text

    Else
        txtWorkSchedule.Text = ""
    End If
End Sub

Private Sub tbF7Sections_Click()
    If tbF7Sections.SelectedItem.Index = 1 Then
        frFReturnToWork.Visible = False
        frHAdditionalWage.Visible = False
        frIWorkSchedule.Visible = False
        frKAdditionalInfo.Visible = False
        frJFilledBy.Visible = False
        frELostTime.Visible = True
        frELostTime.Top = 1880
        frELostTime.Left = 360
    ElseIf tbF7Sections.SelectedItem.Index = 2 Then
        frELostTime.Visible = False
        frHAdditionalWage.Visible = False
        frIWorkSchedule.Visible = False
        frKAdditionalInfo.Visible = False
        frJFilledBy.Visible = False
        frFReturnToWork.Visible = True
        frFReturnToWork.Top = 1880
        frFReturnToWork.Left = 360
    ElseIf tbF7Sections.SelectedItem.Index = 3 Then
        frELostTime.Visible = False
        frFReturnToWork.Visible = False
        frIWorkSchedule.Visible = False
        frKAdditionalInfo.Visible = False
        frJFilledBy.Visible = False
        frHAdditionalWage.Visible = True
        frHAdditionalWage.Top = 1635
        frHAdditionalWage.Left = 360
    ElseIf tbF7Sections.SelectedItem.Index = 4 Then
        frELostTime.Visible = False
        frFReturnToWork.Visible = False
        frHAdditionalWage.Visible = False
        frJFilledBy.Visible = False
        frKAdditionalInfo.Visible = False
        frIWorkSchedule.Visible = True
        frIWorkSchedule.Top = 1880
        frIWorkSchedule.Left = 360
    ElseIf tbF7Sections.SelectedItem.Index = 5 Then
        frELostTime.Visible = False
        frFReturnToWork.Visible = False
        frHAdditionalWage.Visible = False
        frIWorkSchedule.Visible = False
        frKAdditionalInfo.Visible = False
        frJFilledBy.Visible = True
        frJFilledBy.Top = 1880
        frJFilledBy.Left = 360
    ElseIf tbF7Sections.SelectedItem.Index = 6 Then
        frELostTime.Visible = False
        frFReturnToWork.Visible = False
        frHAdditionalWage.Visible = False
        frIWorkSchedule.Visible = False
        frJFilledBy.Visible = False
        frKAdditionalInfo.Visible = True
        frKAdditionalInfo.Top = 1880
        frKAdditionalInfo.Left = 360
    End If
    
End Sub

Private Sub txtAccDecl_Change()
    If txtAccDecl.Text = "A" Then
        optAccDecl(0).Value = True
    ElseIf txtAccDecl.Text = "D" Then
        optAccDecl(1).Value = True
    Else
        optAccDecl(0).Value = False
        optAccDecl(1).Value = False
    End If
End Sub

Private Sub txtBeingPaidYN_Change()
    If txtBeingPaidYN.Text = "" Then
        optBeingPaidYN(0).Value = False
        optBeingPaidYN(1).Value = False
    ElseIf txtBeingPaidYN.Text <> "0" Then
        optBeingPaidYN(0).Value = True
    Else
        optBeingPaidYN(1).Value = True
    End If
End Sub

Private Sub txtConfirmedBy_Change()
    If txtConfirmedBy.Text = "M" Then
        optConfirmedBy(0).Value = True
    ElseIf txtConfirmedBy.Text = "O" Then
        optConfirmedBy(1).Value = True
    Else
        optConfirmedBy(0).Value = False
        optConfirmedBy(1).Value = False
    End If
End Sub

Private Sub txtDiscussed_Change()
    If txtDiscussed.Text = "" Then
        optDiscussed(0).Value = False
        optDiscussed(1).Value = False
    ElseIf txtDiscussed.Text <> "0" Then
        optDiscussed(0).Value = True
    Else
        optDiscussed(1).Value = True
    End If
End Sub

Private Sub txtFilledByName_Change()
    cmbFilledByName.Text = txtFilledByName.Text
End Sub

Private Sub txtFullRegOther_Change()
    If txtFullRegOther.Text = "F" Then
        optFullRegOther(0).Value = True
    ElseIf txtFullRegOther.Text = "O" Then
        optFullRegOther(1).Value = True
    Else
        optFullRegOther(0).Value = False
        optFullRegOther(1).Value = False
    End If
End Sub

Private Sub txtLimitations_Change()
    If txtLimitations.Text = "" Then
        optLimitations(0).Value = False
        optLimitations(1).Value = False
    ElseIf txtLimitations.Text <> "0" Then
        optLimitations(0).Value = True
    Else
        optLimitations(1).Value = True
    End If
End Sub

Private Sub txtLostTime_Change()
    If txtLostTime.Text = "R" Then
        optLostTime(0).Value = True
    ElseIf txtLostTime.Text = "M" Then
        optLostTime(1).Value = True
    ElseIf txtLostTime.Text = "L" Then
        optLostTime(2).Value = True
    Else
        optLostTime(0).Value = False
        optLostTime(1).Value = False
        optLostTime(2).Value = False
    End If
End Sub

Private Sub txtNormFTimeLastWorkAP_Change()
    If txtNormFTimeLastWorkAP.Text = "A" Then
        optNormFTimeLastWorkAP(0).Value = True
    ElseIf txtNormFTimeLastWorkAP.Text = "P" Then
        optNormFTimeLastWorkAP(1).Value = True
    Else
        optNormFTimeLastWorkAP(0).Value = False
        optNormFTimeLastWorkAP(1).Value = False
    End If
End Sub

Private Sub txtNormTTimeLastWorkAP_Change()
    If txtNormTTimeLastWorkAP.Text = "A" Then
        optNormTTimeLastWorkAP(0).Value = True
    ElseIf txtNormTTimeLastWorkAP.Text = "P" Then
        optNormTTimeLastWorkAP(1).Value = True
    Else
        optNormTTimeLastWorkAP(0).Value = False
        optNormTTimeLastWorkAP(1).Value = False
    End If
End Sub

Private Sub txtOffered_Change()
    If txtOffered.Text = "" Then
        optOffered(0).Value = False
        optOffered(1).Value = False
    ElseIf txtOffered.Text <> "0" Then
        optOffered(0).Value = True
    Else
        optOffered(1).Value = True
    End If
End Sub

Private Sub txtOtherEarnings1_Change()
    comOtherEarnings1.Text = txtOtherEarnings1.Text
End Sub

Private Sub txtOtherEarnings2_Change()
    comOtherEarnings2.Text = txtOtherEarnings2.Text
End Sub

Private Sub txtOtherEarnings3_Change()
    comOtherEarnings3.Text = txtOtherEarnings3.Text
End Sub

Private Sub txtOtherEarnings4_Change()
    comOtherEarnings4.Text = txtOtherEarnings4.Text
End Sub

Private Sub txtRegMod_Change()
    If txtRegMod.Text = "R" Then
        optRegMod(0).Value = True
    ElseIf txtRegMod.Text = "M" Then
        optRegMod(1).Value = True
    Else
        optRegMod(0).Value = False
        optRegMod(1).Value = False
    End If
End Sub

Private Sub txtResponsible_Change()
    If txtResponsible.Text = "M" Then
        optResponsible(0).Value = True
    ElseIf txtResponsible.Text = "O" Then
        optResponsible(1).Value = True
    Else
        optResponsible(0).Value = False
        optResponsible(1).Value = False
    End If
End Sub

Private Sub txtTimeLastWorkAP_Change()
    If txtTimeLastWorkAP.Text = "A" Then
        optTimeLastWorkAP(0).Value = True
    ElseIf txtTimeLastWorkAP.Text = "P" Then
        optTimeLastWorkAP(1).Value = True
    Else
        optTimeLastWorkAP(0).Value = False
        optTimeLastWorkAP(1).Value = False
    End If
End Sub

Private Sub txtVacPerctYN_Change()
    If txtVacPerctYN.Text = "" Then
        optVacPerctYN(0).Value = False
        optVacPerctYN(1).Value = False
    ElseIf txtVacPerctYN.Text <> "0" Then
        optVacPerctYN(0).Value = True
    Else
        optVacPerctYN(1).Value = True
    End If
End Sub

Private Sub txtWorkSchedule_Change()
    If txtWorkSchedule.Text = "N" Then
        optWorkSchedule(0).Value = True
    ElseIf txtWorkSchedule.Text = "R" Then
        optWorkSchedule(1).Value = True
    ElseIf txtWorkSchedule.Text = "I" Then
        optWorkSchedule(2).Value = True
    Else
        optWorkSchedule(0).Value = False
        optWorkSchedule(1).Value = False
        optWorkSchedule(2).Value = False
    End If
End Sub

Private Sub Populate_OtherEarnings_ComboBox()
Dim otherEarnlist, X

otherEarnlist = OtherEarningList

X = 1
Do While X > 0
    X = InStr(otherEarnlist, "&")
    If X > 0 Then
        comOtherEarnings1.AddItem Left(otherEarnlist, X - 1)
        comOtherEarnings2.AddItem Left(otherEarnlist, X - 1)
        comOtherEarnings3.AddItem Left(otherEarnlist, X - 1)
        comOtherEarnings4.AddItem Left(otherEarnlist, X - 1)
        otherEarnlist = Mid(otherEarnlist, X + 1)
    Else
        comOtherEarnings1.AddItem otherEarnlist
        comOtherEarnings2.AddItem otherEarnlist
        comOtherEarnings3.AddItem otherEarnlist
        comOtherEarnings4.AddItem otherEarnlist
    End If
Loop

'comOtherEarnings1.Clear
'comOtherEarnings1.AddItem "Bonus"
'comOtherEarnings1.AddItem "Commission"
'comOtherEarnings1.AddItem "Differentials"
'comOtherEarnings1.AddItem "In Lieu %"
'comOtherEarnings1.AddItem "Other (Key In)"
'comOtherEarnings1.AddItem "Premiums"
'comOtherEarnings1.AddItem "Tips"
'
'comOtherEarnings2.Clear
'comOtherEarnings2.AddItem "Bonus"
'comOtherEarnings2.AddItem "Commission"
'comOtherEarnings2.AddItem "Differentials"
'comOtherEarnings2.AddItem "In Lieu %"
'comOtherEarnings2.AddItem "Other (Key In)"
'comOtherEarnings2.AddItem "Premiums"
'comOtherEarnings2.AddItem "Tips"
'
'comOtherEarnings3.Clear
'comOtherEarnings3.AddItem "Bonus"
'comOtherEarnings3.AddItem "Commission"
'comOtherEarnings3.AddItem "Differentials"
'comOtherEarnings3.AddItem "In Lieu %"
'comOtherEarnings3.AddItem "Other (Key In)"
'comOtherEarnings3.AddItem "Premiums"
'comOtherEarnings3.AddItem "Tips"
'
'comOtherEarnings4.Clear
'comOtherEarnings4.AddItem "Bonus"
'comOtherEarnings4.AddItem "Commission"
'comOtherEarnings4.AddItem "Differentials"
'comOtherEarnings4.AddItem "In Lieu %"
'comOtherEarnings4.AddItem "Other (Key In)"
'comOtherEarnings4.AddItem "Premiums"
'comOtherEarnings4.AddItem "Tips"

End Sub

Private Function Refresh_Form7_Data() As Boolean
    Dim xSat As Integer
    Dim xSunDate As Date
    Dim xSatDate As Date

    Refresh_Form7_Data = False
    
    'Update Form 7 sections with the latest information.
    If Not rsDATA.EOF Then
        rsDATA("F7_FED_AMT") = GetEmpData(glbLEE_ID, "ED_TD1DOL", "")
        rsDATA("F7_PROV_AMT") = GetEmpData(glbLEE_ID, "ED_PROVAMT", "")

        'Week 1 - 4 From/To Dates - for Section H-8 and I-C
        'Week is Sun - Sat.
        If IsDate(CVDate(txtIncidentDate.Text)) Then

            'Get which day of the week it is and get the # of days before the last Saturday
            Select Case Weekday(CVDate(txtIncidentDate.Text))
                Case vbMonday
                    xSat = -2
                Case vbTuesday
                    xSat = -3
                Case vbWednesday
                    xSat = -4
                Case vbThursday
                    xSat = -5
                Case vbFriday
                    xSat = -6
                Case vbSaturday
                    xSat = -7
                Case vbSunday
                    xSat = -8
            End Select

            'Week 1 - From
            'Compute the date of the last Saturday
            'xSatDate = DateAdd("d", xSat, CVDate(txtIncidentDate.Text))
            'Date prior to Incident Date
            xSatDate = DateAdd("d", -1, CVDate(txtIncidentDate.Text))

            'Compute the date of last Sunday - 1 week prior
            xSunDate = DateAdd("d", -6, xSatDate)
            rsDATA("F7_OTH_EARN_FROM_WK1") = xSunDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_FWEEK1") = xSunDate

            'Week 1 - To
            rsDATA("F7_OTH_EARN_TO_WK1") = xSatDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_TWEEK1") = xSatDate


            'Week 2 - From
            xSatDate = DateAdd("d", -1, xSunDate)
            xSunDate = DateAdd("d", -6, xSatDate)
            rsDATA("F7_OTH_EARN_FROM_WK2") = xSunDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_FWEEK2") = xSunDate

            'Week 2 - To
            rsDATA("F7_OTH_EARN_TO_WK2") = xSatDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_TWEEK2") = xSatDate


            'Week 3 - From
            xSatDate = DateAdd("d", -1, xSunDate)
            xSunDate = DateAdd("d", -6, xSatDate)
            rsDATA("F7_OTH_EARN_FROM_WK3") = xSunDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_FWEEK3") = xSunDate

            'Week 3 - To
            rsDATA("F7_OTH_EARN_TO_WK3") = xSatDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_TWEEK3") = xSatDate


            'Week 4 - From
            xSatDate = DateAdd("d", -1, xSunDate)
            xSunDate = DateAdd("d", -6, xSatDate)
            rsDATA("F7_OTH_EARN_FROM_WK4") = xSunDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_FWEEK4") = xSunDate

            'Week 4 - To
            rsDATA("F7_OTH_EARN_TO_WK4") = xSatDate
            If optWorkSchedule(2).Value = True Then rsDATA("F7_TWEEK4") = xSatDate
        End If

        rsDATA.Update
        
        Refresh_Form7_Data = True
        
        rsDATA.Requery
        Data1.Recordset.Requery
        
        Call Display_Value
    Else
        Refresh_Form7_Data = False
    End If

End Function
Private Function OtherEarningList() As String
Dim xOtherEarnList As String, mtfFile

'MTF file containing list of Other Earnings
xOtherEarnList = ""
mtfFile = glbIHRREPORTS & "OtherEarningList.MTF"

On Error GoTo ErrorHandler

'Retrieve the Other Earning lists
If File(mtfFile) Then
    Open mtfFile For Input As #1
    Input #1, xOtherEarnList
    Close #1
End If

ResumeHere:

'All four Other Earning combobox should have the same set of list.
'Add any new values to the existing list
If InStr(xOtherEarnList, comOtherEarnings1) = 0 And comOtherEarnings1 <> "" Then
    xOtherEarnList = xOtherEarnList & "&" & comOtherEarnings1
End If

If InStr(xOtherEarnList, comOtherEarnings2) = 0 And comOtherEarnings2 <> "" Then
    xOtherEarnList = xOtherEarnList & "&" & comOtherEarnings2
End If

If InStr(xOtherEarnList, comOtherEarnings3) = 0 And comOtherEarnings3 <> "" Then
    xOtherEarnList = xOtherEarnList & "&" & comOtherEarnings3
End If

If InStr(xOtherEarnList, comOtherEarnings4) = 0 And comOtherEarnings4 <> "" Then
    xOtherEarnList = xOtherEarnList & "&" & comOtherEarnings4
End If

'Resave the list of Other Earning values to MTF file as it may contain new values input by the user on
'the combo box.
Open mtfFile For Output As #1
Print #1, xOtherEarnList
Close #1

OtherEarningList = xOtherEarnList

Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt OtherEarningList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Other Earning List"
    Kill mtfFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub Populate_FilledBy_Names()
    Dim rsFilledBy As New ADODB.Recordset
    Dim SQLQ As String
    
    cmbFilledByName.Clear
    
    SQLQ = "SELECT P7_NAME FROM HR_OHS_PERSON_COMPLTG_F7 WHERE P7_INACTIVE <> 1 ORDER BY P7_NAME"
    rsFilledBy.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Do While Not rsFilledBy.EOF
        cmbFilledByName.AddItem rsFilledBy("P7_NAME")
    
        rsFilledBy.MoveNext
    Loop
    rsFilledBy.Close
    Set rsFilledBy = Nothing
End Sub

Private Sub Get_FilledBy_Details()
    Dim rsFilledBy As New ADODB.Recordset
    Dim SQLQ As String
       
    SQLQ = "SELECT * FROM HR_OHS_PERSON_COMPLTG_F7 WHERE P7_INACTIVE <> 1 AND P7_NAME = '" & cmbFilledByName.Text & "'"
    rsFilledBy.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsFilledBy.EOF Then
        txtFilledByTitle.Text = rsFilledBy("P7_TITLE")
        medFilledByTelephone.Text = rsFilledBy("P7_PHONE")
    End If
    rsFilledBy.Close
    Set rsFilledBy = Nothing
End Sub
