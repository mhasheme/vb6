VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEContEmpDemo 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Contractor"
   ClientHeight    =   9645
   ClientLeft      =   270
   ClientTop       =   1305
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
   FontTransparent =   0   'False
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9645
   ScaleWidth      =   11670
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frContractual 
      Caption         =   "Contractual Data"
      Height          =   2535
      Left            =   120
      TabIndex        =   37
      Top             =   6120
      Width           =   11415
      Begin VB.TextBox txtReptAuthority 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   54
         Tag             =   "00-Employee Number of individual's supervisor"
         Top             =   1110
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox comPayPer 
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
         Left            =   4410
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "01-Choose annum or hour"
         Top             =   1470
         Width           =   1190
      End
      Begin INFOHR_Controls.DateLookup dlpStartDate 
         DataSource      =   " "
         Height          =   285
         Left            =   1515
         TabIndex        =   17
         Tag             =   "40-Start Date"
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpEndDate 
         DataSource      =   " "
         Height          =   285
         Left            =   7320
         TabIndex        =   18
         Tag             =   "40-End Date"
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         Height          =   285
         Left            =   1515
         TabIndex        =   19
         Tag             =   "01-Position code"
         Top             =   720
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
         Height          =   285
         Index           =   0
         Left            =   1515
         TabIndex        =   20
         Tag             =   "10-Employee Number of individual's supervisor"
         Top             =   1110
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin MSMask.MaskEdBox medHours 
         Height          =   285
         Index           =   0
         Left            =   1830
         TabIndex        =   24
         Tag             =   "10-Usual working hours per day"
         Top             =   1875
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
      Begin MSMask.MaskEdBox medHours 
         Height          =   285
         Index           =   1
         Left            =   4410
         TabIndex        =   25
         Tag             =   "10- Number of hours in work week"
         Top             =   1875
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
      Begin MSMask.MaskEdBox medHours 
         Height          =   285
         Index           =   2
         Left            =   7635
         TabIndex        =   26
         Tag             =   "10-Usual working hours per pay period"
         Top             =   1875
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   9
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
      Begin MSMask.MaskEdBox medsalary 
         Height          =   285
         Left            =   1830
         TabIndex        =   21
         Tag             =   "21-Enter salary"
         Top             =   1485
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   7320
         TabIndex        =   23
         Tag             =   "00-Currency Indicator - Code"
         Top             =   1485
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFCI"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   5640
         TabIndex        =   65
         Top             =   405
         Width           =   675
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   120
         TabIndex        =   64
         Top             =   405
         Width           =   885
      End
      Begin VB.Label lblPosTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   63
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblReptAuth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   1155
         Width           =   1395
      End
      Begin VB.Label lblHrsDay 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Day"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label lblHrsWeek 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Week"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2925
         TabIndex        =   60
         Top             =   1920
         Width           =   1170
      End
      Begin VB.Label lblHrsPayPeriod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours/Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5490
         TabIndex        =   59
         Top             =   1920
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Per"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   3570
         TabIndex        =   58
         Top             =   1530
         Width           =   525
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   120
         TabIndex        =   57
         Top             =   1530
         Width           =   1380
      End
      Begin VB.Label lblCurrencyIndicator 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
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
         Left            =   6480
         TabIndex        =   56
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SalCode"
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5160
         TabIndex        =   55
         Top             =   1530
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.Frame frOrganizational 
      Caption         =   "Organizational Data"
      Height          =   1695
      Left            =   120
      TabIndex        =   36
      Top             =   4320
      Width           =   11415
      Begin VB.TextBox txtIPHONE 
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
         Left            =   7515
         MaxLength       =   10
         TabIndex        =   16
         Tag             =   "00-Internal Telephone Extension "
         Top             =   1080
         Width           =   1305
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1515
         TabIndex        =   13
         Tag             =   "00-Department"
         Top             =   1080
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1515
         TabIndex        =   11
         TabStop         =   0   'False
         Tag             =   "00-Division"
         Top             =   360
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
         Enabled         =   0   'False
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1515
         TabIndex        =   12
         Tag             =   "00-Location - Code"
         Top             =   720
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   7200
         TabIndex        =   14
         Tag             =   "00-Section"
         Top             =   360
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   7200
         TabIndex        =   15
         Tag             =   "00-Region"
         Top             =   720
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   53
         Top             =   1125
         Width           =   990
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   52
         Top             =   405
         Width           =   690
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Section"
         Height          =   195
         Index           =   9
         Left            =   5640
         TabIndex        =   51
         Top             =   405
         Width           =   660
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   5640
         TabIndex        =   50
         Top             =   765
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   49
         Top             =   765
         Width           =   750
      End
      Begin VB.Label lblIPhone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Phone Extension"
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
         Left            =   5640
         TabIndex        =   48
         Top             =   1125
         Width           =   1770
      End
   End
   Begin VB.Frame frPersonal 
      Caption         =   "Personal Data"
      Height          =   3495
      Left            =   120
      TabIndex        =   35
      Top             =   720
      Width           =   11415
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
         Height          =   285
         Left            =   1830
         MaxLength       =   60
         TabIndex        =   10
         Tag             =   "00-Email Address"
         Top             =   2920
         Width           =   4260
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
         Left            =   6840
         TabIndex        =   7
         Tag             =   "00-Country"
         Top             =   2175
         Width           =   1320
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   8190
         MaxLength       =   10
         TabIndex        =   38
         Tag             =   "01-Country"
         Top             =   2190
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSurname 
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
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   0
         Tag             =   "01-Surname"
         Top             =   360
         Width           =   4180
      End
      Begin VB.TextBox txtFName 
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
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "01-First or Given Name"
         Top             =   720
         Width           =   4180
      End
      Begin VB.TextBox txtAdd1 
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
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   2
         Tag             =   "01-First Line in Address"
         Top             =   1095
         Width           =   4180
      End
      Begin VB.TextBox txtAdd2 
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
         Height          =   285
         Left            =   1830
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "00-Second Line in Address"
         Top             =   1455
         Width           =   4180
      End
      Begin VB.TextBox txtCity 
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
         Height          =   285
         Left            =   1830
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "01-City"
         Top             =   1815
         Width           =   2895
      End
      Begin MSMask.MaskEdBox medPCode 
         Height          =   285
         Left            =   1830
         TabIndex        =   6
         Tag             =   "40-Postal Code"
         Top             =   2190
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
         Height          =   285
         Left            =   1830
         TabIndex        =   8
         Tag             =   "11-Telephone Number"
         Top             =   2550
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
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   6525
         TabIndex        =   5
         Tag             =   "31-Province - Code"
         Top             =   1815
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin MSMask.MaskEdBox medCellPhone 
         Height          =   285
         Left            =   6840
         TabIndex        =   9
         Tag             =   "10-Cellular Telephone Number"
         Top             =   2550
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
      Begin VB.Label lblEmail 
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
         Left            =   120
         TabIndex        =   66
         Top             =   2965
         Width           =   1455
      End
      Begin VB.Label lblTitle 
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
         Index           =   20
         Left            =   5040
         TabIndex        =   47
         Top             =   2595
         Width           =   1020
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Surname"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   405
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   45
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   44
         Top             =   1140
         Width           =   690
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   43
         Top             =   1860
         Width           =   330
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   120
         TabIndex        =   42
         Top             =   2595
         Width           =   915
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Province"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   5040
         TabIndex        =   41
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Postal Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   40
         Top             =   2235
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   5040
         TabIndex        =   39
         Top             =   2235
         Width           =   1260
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   33
      Top             =   8985
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Save && Close"
         Height          =   375
         Left            =   3773
         TabIndex        =   27
         Tag             =   "Save the changes and exit this screen"
         Top             =   120
         Width           =   1845
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   6053
         TabIndex        =   28
         Tag             =   "Cancel the changes and exit this screen"
         Top             =   120
         Width           =   1845
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11670
      _Version        =   65536
      _ExtentX        =   20585
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
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataField       =   "ED_EMPNBR"
         DataSource      =   "Data1"
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   8880
         TabIndex        =   34
         Top             =   165
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   32
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
         TabIndex        =   30
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
         TabIndex        =   31
         Top             =   135
         Width           =   1245
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      Caption         =   ""
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
Attribute VB_Name = "frmEContEmpDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsORG As New ADODB.Recordset
Dim RDept, RGLNum
Dim OHOMELINE, OHOMESHIFT, OHOMEOPRTNBR, oHOMEWRKCNT
Dim xUpdateable
Dim fglbJobList As String
Dim fglbNewEE As Integer
Dim oldEEId
Dim SavDiv, SavDept
Dim flagFrmLoad As Boolean

Private Function chk_ContEmployee()
Dim X%
Dim Msg$, Response%

chk_ContEmployee = False

'Check Employee Number on Save
If fglbNewEE = True Then
    If Not chk_EMPNBR Then Exit Function
End If


If (UCase(comCountry) = "CANADA" Or UCase(comCountry) = "U.S.A.") Then
    If Len(txtSurname.Text) > 0 Then 'Ticket #14154
        If Len(InvalidCharInStr(txtSurname.Text, glbWFCNameChars)) > 0 Then
            MsgBox "Invalid character '" & InvalidCharInStr(txtSurname.Text, glbWFCNameChars) & "' in name field. "
            txtSurname.SetFocus
            Exit Function
        End If
    End If
    If Len(txtFName.Text) > 0 Then 'Ticket #14154
        If Len(InvalidCharInStr(txtFName.Text, glbWFCNameChars)) > 0 Then
            MsgBox "Invalid character '" & InvalidCharInStr(txtFName.Text, glbWFCNameChars) & "' in name field. "
            txtFName.SetFocus
            Exit Function
        End If
    End If
End If
    
If Not (UCase(comCountry) = "CHINA") Then
    Msg = "Do not use ALL CAPS for name and address fields." & Chr(10) & "Please enter data using Proper/Title Case Only."
    If Len(txtSurname.Text) > 0 Then
        If AllCapitalString(txtSurname.Text) Then
            MsgBox Msg: txtSurname.SetFocus: Exit Function
        End If
    End If
    If Len(txtFName.Text) > 0 Then
        If AllCapitalString(txtFName.Text) Then
            MsgBox Msg: txtFName.SetFocus: Exit Function
        End If
    End If
    If Len(txtAdd1.Text) > 0 Then
        If AllCapitalString(txtAdd1.Text, "RR") Then
            MsgBox Msg: txtAdd1.SetFocus: Exit Function
        End If
    End If
    If Len(txtAdd2.Text) > 0 Then
        If AllCapitalString(txtAdd2.Text, "RR") Then
            MsgBox Msg: txtAdd2.SetFocus: Exit Function
        End If
    End If
    If Len(txtCity.Text) > 0 Then
        If AllCapitalString(txtCity.Text) Then
            MsgBox Msg: txtCity.SetFocus: Exit Function
        End If
    End If
End If

If Len(txtSurname) < 1 Then
    MsgBox "Surname is a required field"
    txtSurname.SetFocus
    Exit Function
End If

If Len(txtFName) < 1 Then
    MsgBox lStr("First Name is a required field")
    txtFName.SetFocus
    Exit Function
End If

If (Not gSec_Show_ADDRESS) And (Len(txtAdd1) < 1) Then
    MsgBox "First Address Line is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'." & vbCrLf & vbCrLf & "Cancelling the changes."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(txtAdd1) < 1 Then
        MsgBox "First Address Line is a required field"
        txtAdd1.SetFocus
        Exit Function
    End If
End If

If (Not gSec_Show_ADDRESS) And (Len(txtCity) < 1) Then
    MsgBox "City is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(txtCity) < 1 Then
        MsgBox "City is a required field"
        txtCity.SetFocus
        Exit Function
    End If
End If

If (Not gSec_Show_ADDRESS) And ((Len(clpProv.Text) < 1) Or (clpProv.Caption = "Unassigned")) Then
    MsgBox "Province is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(clpProv.Text) < 1 Then
        MsgBox "Province is a required field"
        clpProv.SetFocus
        Exit Function
    Else
        If clpProv.Caption = "Unassigned" Then
            MsgBox "Invalid Province"
            clpProv.SetFocus
            Exit Function
        End If
    End If
End If

If (Not gSec_Show_ADDRESS) And (Len(medPCode) < 1) Then
    MsgBox "Postal/Zip Code is a required field." & vbCrLf & "To allow the data entry for 'Address', please check the Security Setup for 'Address'."
    Exit Function
ElseIf gSec_Show_ADDRESS Then
    If Len(medPCode) < 1 Then
        If comCountry = "U.S.A." Or comCountry = "MEXICO" Then
            MsgBox "Zip Code is a required field"
            medPCode.SetFocus
            Exit Function
        End If
        If comCountry = "CANADA" Then
            MsgBox "Postal Code is a required field"
            medPCode.SetFocus
            Exit Function
        End If
    End If
End If

If Len(comCountry.Text) = 0 Then
    MsgBox ("Country is a required field")
    comCountry.SetFocus
    Exit Function
End If

If Len(medTelephone) < 1 Then
    MsgBox "Telephone Number is a required field"
    medTelephone.SetFocus
    Exit Function
End If

For X = 1 To 4
    If X = 3 Then GoTo nextcode
    If Len(clpCode(X).Text) > 0 And clpCode(X).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X).SetFocus
        Exit Function
    End If
nextcode:
Next X

If Len(clpCode(1).Text) = 0 Then
    MsgBox lStr("Location is a required field")
    clpCode(1).SetFocus
    Exit Function
End If

If Len(clpDept.Text) < 1 Then
    MsgBox lStr("Department is a required field")
    clpDept.SetFocus
    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        MsgBox "Department Code must be valid"
        clpDept.SetFocus
        Exit Function
    End If
End If

If Len(clpCode(4).Text) = 0 Then
    MsgBox lStr("Section is a required field")
    clpCode(4).SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) < 1 Then
    MsgBox lStr("Region is a required field")
    clpCode(2).SetFocus
    Exit Function
End If

If Len(dlpStartDate.Text) < 1 Then
    MsgBox "Position Start Date must be entered"
    dlpStartDate.SetFocus
    Exit Function
ElseIf Not IsDate(dlpStartDate.Text) Then
    MsgBox "Position Start Date is not a valid date"
    dlpStartDate.SetFocus
    Exit Function
End If

If Len(dlpEndDate.Text) > 0 Then
    If Not IsDate(dlpEndDate.Text) Then
        MsgBox "End Date is not a valid date"
        dlpEndDate.SetFocus
        Exit Function
    End If

    If IsDate(dlpEndDate.Text) Then
        If DateDiff("d", dlpEndDate.Text, dlpStartDate.Text) > 0 Then
            MsgBox "End Date must be later than Start Date"
            dlpEndDate.SetFocus
            Exit Function
        End If
    End If
End If

If Len(clpJob.Text) <= 0 Then
    MsgBox "Position Code is required"
    clpJob.SetFocus
    Exit Function
Else
    If clpJob.Caption = "Unassigned" Then
        MsgBox "Position Code is required"
        clpJob.SetFocus
        Exit Function
    End If
End If

If IsInactivePos(clpJob.Text) Then
    MsgBox "'" & clpJob.Text & "' is Inactive Position Code. Please contact Corporate Total Rewards to review this Position Requirement."
    Exit Function
End If

If IsMissingBudPos(clpJob.Text) Then
    MsgBox "Please contact the info:HR corporate administrator to have them create the Budgeted Position Master for '" & clpJob.Text & "' "
    Exit Function
End If

If (glbUNION = "NONE" Or glbUNION = "EXEC") Then 'Salary employee only
    If Len(txtReptAuthority(0).Text) > 0 Then
        If IsRept1PosNotMatchPosMaster(txtReptAuthority(0).Text, clpJob.Text) Then
            glbMsgCustomVal = 11
            frmMsgDialog.Show 1
            'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
            If glbMsgCustomVal = 2 Then 'If <<Cancel>> is checked, undo the change.
                'Call cmdCancel_Click
                txtReptAuthority(0).Text = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
                Exit Function
            End If
        End If
    End If
End If

For X% = 0 To 0 '2
    If elpReptAuthShow(X%) = "0" Then elpReptAuthShow(X%) = ""
    If Len(elpReptAuthShow(X%)) > 0 Then
        If elpReptAuthShow(X%).Caption = "Unassigned" Then
            MsgBox "Rept. Authority Employee # not valid. Check Employee # and re-enter!"
            elpReptAuthShow(X%).SetFocus
            Exit Function
        End If
    End If
Next

If Len(elpReptAuthShow(0)) = 0 Then
    MsgBox "Rept. Authority 1 is required."
    elpReptAuthShow(0).SetFocus
    Exit Function
End If

If Len(medsalary) < 1 Then
    MsgBox "Salary is required"
    medsalary.SetFocus
    Exit Function
End If
If medsalary <= 0 Then
    MsgBox "Salary is required"
    medsalary.SetFocus
    Exit Function
End If

If comPayPer.Text = "" Then
    MsgBox "Per cannot be blank"
    comPayPer.SetFocus
    Exit Function
End If


If Not IsNumeric(medHours(0)) Then
    MsgBox "Hours/Day is required"
    medHours(0).SetFocus
    Exit Function
End If
If Not IsNumeric(medHours(1)) Then
    MsgBox "Hours/Week is required"
    medHours(1).SetFocus
    Exit Function
End If
If Not IsNumeric(medHours(2)) Then
    MsgBox "Hours/Per Period is required"
    medHours(2).SetFocus
    Exit Function
End If

chk_ContEmployee = True

End Function

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
            
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = lblEEID
    'rsTA("AU_Payroll_ID") = lblEEID     'Jerry said CONP employees do not go to Payroll so we are not going to populate it.
            
    rsTA("AU_SURNAME") = txtSurname
    rsTA("AU_FNAME") = txtFName
    rsTA("AU_ADDR1") = txtAdd1
    If Trim$(txtAdd2) <> "" Then rsTA("AU_ADDR2") = txtAdd2
    rsTA("AU_CITY") = txtCity
    rsTA("AU_PROV") = clpProv.Text
    rsTA("AU_PROVRES") = xProvNbr
    rsTA("AU_PROVEMP") = clpProv.Text
    rsTA("AU_PCODE") = medPCode
    rsTA("AU_COUNTRY") = txtCountry
    rsTA("AU_SIN") = "999999999"
    rsTA("AU_DOB") = CVDate("01/01/1995")
    rsTA("AU_MSTAT") = "O"
    rsTA("AU_SEX") = "N"
    rsTA("AU_PHONE") = medTelephone
    rsTA("AU_PAGENBR") = medCellPhone.Text
    
    rsTA("AU_DIV") = clpDiv.Text
    rsTA("AU_DIVUPL") = clpDiv.Text
    rsTA("AU_DEPTNO") = clpDept.Text
    rsTA("AU_DEPT_GL") = "0000"
    rsTA("AU_DEPTEDATE") = dlpStartDate.Text
    rsTA("AU_DIVEDATE") = dlpStartDate.Text
    
    rsTA("AU_LOC") = clpCode(1).Text
    rsTA("AU_ADMINBY") = "LMER"
    rsTA("AU_REGION") = clpCode(2).Text
    rsTA("AU_SECTION") = clpCode(4).Text
    
    rsTA("AU_NEWEMP") = "Y"
End If

If xADD Then
    'Status/Dates
    rsTA("AU_INTEL") = txtIPHONE
    rsTA("AU_EMAIL") = txtEmail
    
    rsTA("AU_EMP") = "CONP"
    rsTA("AU_SFDATE") = dlpStartDate.Text
    rsTA("AU_PT") = "CE"
    rsTA("AU_PTUPL") = "CE"
    rsTA("AU_PTEDATE") = dlpStartDate.Text
    rsTA("AU_ORG") = "CE"
    rsTA("AU_DOH") = dlpStartDate.Text
    
    'Position
    rsTA("AU_JOB") = clpJob.Text
    rsTA("AU_SJDATE") = dlpStartDate.Text
    rsTA("AU_JREASON") = "CE"
    rsTA("AU_DHRS") = medHours(0)
    rsTA("AU_WHRS") = medHours(1)
    rsTA("AU_PHRS") = medHours(2)
        
    'Salary
    rsTA("AU_SALARY") = medsalary
    rsTA("AU_SALCD") = "H"
    rsTA("AU_SEDATE") = dlpStartDate.Text
    rsTA("AU_SREASON") = "CE"
    rsTA("AU_PAYP") = "T"
    
    rsTA("AU_LDATE") = Date
    If IsDate(dlpStartDate.Text) Then
        If CVDate(dlpStartDate.Text) > Date Then
            rsTA("AU_LDATE") = dlpStartDate.Text
        End If
    End If
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = Actn
    rsTA.Update
        
End If

AUDITDEMO = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
'Resume Next

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING CONTRACT EMP AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '18June99 js

Resume Next

End Function

Private Sub clpCode_Change(Index As Integer)
    If Index = 4 Then
        clpCode(0).Text = getWFCCurrencyIndi(clpCode(4).Text)
    End If
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpDept_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpDiv_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpJob_Change()
    If Len(txtReptAuthority(0).Text) = 0 And Len(clpJob.Text) > 0 Then
        txtReptAuthority(0).Text = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
    End If
End Sub

Private Sub clpJob_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub clpJob_LostFocus()
    If Len(txtReptAuthority(0).Text) = 0 And Len(clpJob.Text) > 0 Then
        txtReptAuthority(0).Text = GetReportingAuth1EmpNoBasePosMaster(clpJob.Text)
    End If
End Sub

Private Sub clpProv_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Sub cmdCancel_Click()
Dim X

If fglbNewEE = True Then
    fglbNewEE = False
    
    'Ticket #29660 - Contract Employee - User cancelled the Add process so undo the last # allocated.
    If glbWFC And glbWFCContractEmployee Then
        Call WFC_CancelContractEmployeeNo
    End If
    
    'Reset the global Employee # to the employee # the user had selected last
    glbLEE_ID = oldEEId
    
    If glbOnTop = "FRMECONTEMPDEMO" Then glbOnTop = ""
    
    Unload Me
    
    Exit Sub
End If

If glbOnTop = "FRMECONTEMPDEMO" Then glbOnTop = ""

Unload Me

'Call ST_UPD_MODE(True)

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '21June99 js

End Sub

Public Sub cmdClose_Click()
Unload Me
If glbOnTop = "FRMECONTEMPDEMO" Then glbOnTop = ""
End Sub

Public Sub cmdNew_Click()
Dim X%
Dim intRet
Dim sNewEmpNbr As String
Dim FEENum  As Long
Dim Msg As String, Msg1 As String, Title As String
Dim SQLQ As String

If Not modECountChk() Then
    MsgBox "You have reached the maximum number of employees for your license"
    Exit Sub
End If

MDIMain.MainToolBar.ButtonS(8).Enabled = True
MDIMain.MainToolBar.ButtonS(9).Enabled = True

fglbNewEE = True
oldEEId = glbLEE_ID

glbWFCContractEmployee = True

Do
    sNewEmpNbr = GetNewEmpnbr()
    
    If Len(sNewEmpNbr) > 0 Then
        If Not IsNumeric(sNewEmpNbr) Then
            Msg = Msg1 & "Sorry, must be numeric."
            GoTo NEW_NG
        End If
        If Len(sNewEmpNbr) > 9 Then
            Msg = Msg1 & "Number must be between 1 and 999999999"
            GoTo NEW_NG
        End If
        FEENum = CLng(sNewEmpNbr)
        If FEENum < 1 Or FEENum > 999999999 Then
            Msg = Msg1 & "Number must be between 1 and 999999999"
            GoTo NEW_NG
        End If
        If glbWFC Then
            If Len(sNewEmpNbr) <> 8 Then
                Msg = Msg1 & "Employee Number is not valid format" & Chr(10)
                Msg = Msg & "It should be Division Number + 4 digit Employee Number"
                GoTo NEW_NG
            End If
        End If
        
        Dim rsEmp As New ADODB.Recordset
        SQLQ = "Select ED_EMPNBR from HREMP"
        SQLQ = SQLQ & " where ED_EMPNBR = " & FEENum & ""
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If rsEmp.BOF And rsEmp.EOF Then
            Call SET_UP_MODE
            
            glbLEE_ID = CLng(sNewEmpNbr)
            
            rsEmp.Close
            
            Exit Do
        Else
            Msg = "Sorry, Employee # " & ShowEmpnbr(sNewEmpNbr)
            Msg = Msg & Chr(10) & rsEmp("ED_SURNAME")
            Msg = Msg & Chr(10) & "Already exists."
            rsEmp.Close
            GoTo NEW_NG
        End If
    Else
        MsgBox "Add of new Contractor Aborted"
        glbWFCContractEmployee = False
        fglbNewEE = False
        glbLEE_ID = 0
                
        MDIMain.panHelp(0).Caption = "Select function from the menu."   'Jaddy 10/22/99
        
        Unload Me
        
        Exit Do
    End If

NEW_NG:
        MsgBox Msg, , Title
Loop
    
If fglbNewEE = False Then Exit Sub

'Generate Employee #
lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
lblEEName.Caption = "New Contractor"

clpDiv.Text = glbTrsDIV
clpDept.Text = glbTrsDept

'Get other field values from the Division selected
Call SetOtherFieldsFromDiv(clpDiv.Text)

'Get Currency
clpCode(0).Text = getWFCCurrencyIndi(clpCode(4).Text)

'comCountry = GetCountryFromDiv(clpDiv.Text)
'If IsNull(glbCountry) Then
'    comCountry = "CANADA"
'Else
'    comCountry = glbCountry
'End If

Call SetCountries


MDIMain.MainToolBar.ButtonS(8).Visible = True
MDIMain.MainToolBar.ButtonS(9).Visible = True
MDIMain.MainToolBar.ButtonS(8).Enabled = True
MDIMain.MainToolBar.ButtonS(9).Enabled = True
MDIMain.MainToolBar.ButtonS(1).Enabled = True

End Sub

Private Sub cmdCancel_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub cmdOK_Click()
Dim rsLDiv As New ADODB.Recordset 'Ticket #30491 Franks 09/07/2017
Dim X%

    If Not IsNumeric(medHours(0)) Then medHours(0) = 0
    If Not IsNumeric(medHours(1)) Then medHours(1) = 0
    If Not IsNumeric(medHours(2)) Then medHours(2) = 0

    If Not chk_ContEmployee() Then Exit Sub
    
    If fglbNewEE Then
        'Save records in various tables based on the data entry
    
    
        'Update Employee History
        If Len(Trim(clpDept.Text)) > 0 Then
            If Not EmpHisCalc(2, glbLEE_ID, clpDept, "", "", "", "", "", "", dlpStartDate.Text) Then MsgBox "EMPHIS Error "
        End If
        If Len(Trim(clpDiv.Text)) > 0 Then
            If Not EmpHisCalc(2, glbLEE_ID, "", clpDiv, "", "", "", "", "", dlpStartDate.Text) Then MsgBox "EMPHIS Error "
        End If
        If Len(Trim(clpCode(1))) > 0 Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpStartDate.Text, "LOC", clpCode(1)) Then MsgBox "EMPHIS Error "
        If Len(Trim(clpCode(2))) > 0 Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpStartDate.Text, "REGION", clpCode(2)) Then MsgBox "EMPHIS Error "
        If Len(Trim(clpCode(4))) > 0 Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpStartDate.Text, "SECTION", clpCode(4)) Then MsgBox "EMPHIS Error "
        
        'Default values
        If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", dlpStartDate.Text, "ADMINBY", "LMER") Then MsgBox "EMPHIS Error "
        If Not EmpHisCalc(2, glbLEE_ID, "", "", "CONP", "", "", "", "", dlpStartDate.Text, , , dlpStartDate.Text) Then MsgBox "EMPHIS Error"
        If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "CE", "", "", "", dlpStartDate.Text) Then MsgBox "EMPHIS Error"
        If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "CE", "", "", dlpStartDate.Text) Then MsgBox "EMPHIS Error"
        If Not EmpHisCalc(6, glbLEE_ID, "", "", "", "", "", "", "", Date, , "O") Then MsgBox "EMPHIS Error "
        
        
        Dim rsHREmp As New ADODB.Recordset
        Dim rsEmpOther As New ADODB.Recordset
        Dim rsEmpJob As New ADODB.Recordset
        Dim rsEmpSal As New ADODB.Recordset
        Dim SQLQ As String
        
        'Create a New Hire in HREMP
        SQLQ = "SELECT * FROM HREMP WHERE 1 = 2"
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsHREmp.AddNew
        rsHREmp("ED_COMPNO") = "001"
        rsHREmp("ED_EMPNBR") = glbLEE_ID
        'rsHREmp("ED_PAYROLL_ID") = glbLEE_ID    'Jerry said CONP employees do not go to Payroll so we are not going to populate it.
        rsHREmp("ED_SURNAME") = txtSurname.Text
        rsHREmp("ED_FNAME") = txtFName.Text
        rsHREmp("ED_ADDR1") = txtAdd1.Text
        rsHREmp("ED_ADDR2") = txtAdd2.Text
        rsHREmp("ED_CITY") = txtCity.Text
        rsHREmp("ED_PCODE") = medPCode
        rsHREmp("ED_PROV") = clpProv.Text
        rsHREmp("ED_PROVEMP") = clpProv.Text
        
        rsHREmp("ED_COUNTRY") = txtCountry
        rsHREmp("ED_WORKCOUNTRY") = txtCountry
        rsHREmp("ED_DOB") = CVDate("01/01/1995")
        rsHREmp("ED_SIN") = "999999999"
        rsHREmp("ED_DOH") = dlpStartDate.Text
        rsHREmp("ED_MSTAT") = "O"
        rsHREmp("ED_SEX") = "N"
        rsHREmp("ED_PHONE") = medTelephone
        rsHREmp("ED_PAGENBR") = medCellPhone
        
        rsHREmp("ED_DIV") = clpDiv.Text
        rsHREmp("ED_LOC") = clpCode(1).Text
        rsHREmp("ED_DEPTNO") = clpDept.Text
        rsHREmp("ED_SECTION") = clpCode(4).Text
        rsHREmp("ED_REGION") = clpCode(2).Text
        rsHREmp("ED_INTEL") = txtIPHONE.Text
        
        rsHREmp("ED_DEPTEDATE") = dlpStartDate.Text
        rsHREmp("ED_DIVEDATE") = dlpStartDate.Text
        rsHREmp("ED_GLNO") = "0000"
        rsHREmp("ED_BONUSDEPT") = "000000"
        'rsHREmp("ED_ADMINBY") = "LMER"
        SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDiv.Text & "'"
        rsLDiv.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsLDiv.EOF Then
            rsHREmp("ED_ADMINBY") = rsLDiv("DV_ADMINBY")
        End If
        rsLDiv.Close
        'Ticket #30491 Franks 09/07/2017 - end
        rsHREmp("ED_ORGT1") = "002"
        
        rsHREmp("ED_EMP") = "CONP"
        rsHREmp("ED_SFDATE") = dlpStartDate.Text
        rsHREmp("ED_PT") = "CE"
        rsHREmp("ED_PTEDATE") = dlpStartDate.Text
        rsHREmp("ED_ORG") = "CE"
        
        glbUNION = rsHREmp("ED_ORG")
        
        rsHREmp("ED_EMAIL") = txtEmail.Text
        rsHREmp("ED_VADIM2") = "NA"
        rsHREmp("ED_LDATE") = Date
        rsHREmp("ED_LTIME") = Time$
        rsHREmp("ED_LUSER") = glbUserID
        rsHREmp.Update
        rsHREmp.Close
        Set rsHREmp = Nothing
        
        'HREMP_OTHER
        SQLQ = "SELECT * FROM HREMP_OTHER"
        SQLQ = SQLQ & " WHERE ER_EMPNBR = " & glbLEE_ID
        rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEmpOther.EOF Then
            rsEmpOther.AddNew
            rsEmpOther("ER_COMPNO") = "001"
            rsEmpOther("ER_EMPNBR") = glbLEE_ID
            rsEmpOther("ER_NETWORKLOGIN") = Left(NetworkLoginGenerator, 40)
            rsEmpOther("ER_VENDORNO") = "n/a"
            rsEmpOther("ER_LDATE") = Date
            rsEmpOther("ER_LTIME") = Time$
            rsEmpOther("ER_LUSER") = glbUserID
            rsEmpOther.Update
        End If
        rsEmpOther.Close
        Set rsEmpOther = Nothing
        
        'Create a new record in the HR_JOB_HISTORY table for Contract Employees
        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE 1 = 2"
        rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEmpJob.AddNew
        rsEmpJob("JH_COMPNO") = "001"
        rsEmpJob("JH_EMPNBR") = glbLEE_ID
        rsEmpJob("JH_SDATE") = dlpStartDate.Text
        If IsDate(dlpEndDate.Text) Then rsEmpJob("JH_ENDDATE") = dlpEndDate.Text
        rsEmpJob("JH_CURRENT") = "1"
        rsEmpJob("JH_JOB") = clpJob.Text
        
        rsEmpJob("JH_REPTAU") = elpReptAuthShow(0)
        If IsDate(dlpStartDate.Text) Then rsEmpJob("JH_EDATEREPT1") = dlpStartDate.Text
        rsEmpJob("JH_DHRS") = medHours(0).Text
        rsEmpJob("JH_WHRS") = medHours(1).Text
        rsEmpJob("JH_PHRS") = medHours(2).Text
        rsEmpJob("JH_SHIFT") = "NS"
        rsEmpJob("JH_JREASON") = "CE"
        rsEmpJob("JH_FTENUM") = 1
        rsEmpJob("JH_FTEHRS") = 0 ' 2080
        
        rsEmpJob("JH_DIV") = clpDiv.Text
        rsEmpJob("JH_DEPTNO") = clpDept.Text
        rsEmpJob("JH_EMP") = "CONP"
        rsEmpJob("JH_ORG") = "CE"
        rsEmpJob("JH_PT") = "CE"
        rsEmpJob("JH_SECTION") = clpCode(4).Text
        
        rsEmpJob("JH_LDATE") = Date
        rsEmpJob("JH_LTIME") = Time$
        rsEmpJob("JH_LUSER") = glbUserID
        rsEmpJob.Update
        rsEmpJob.Close
        Set rsEmpJob = Nothing
        
        
        'Create a new record in the HR_JOB_HISTORY table for Contract Employees
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE 1 = 2"
        rsEmpSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEmpSal.AddNew
        rsEmpSal("SH_COMPNO") = "001"
        rsEmpSal("SH_EMPNBR") = glbLEE_ID
        rsEmpSal("SH_JOB") = clpJob.Text
        rsEmpSal("SH_SDATE") = dlpStartDate.Text
        rsEmpSal("SH_WHRS") = medHours(1).Text
        rsEmpSal("SH_CURRENT") = "1"
        rsEmpSal("SH_SREAS1") = "CE"
        rsEmpSal("SH_SALPC1") = 1
        rsEmpSal("SH_SALCHG1") = medsalary.Text
        rsEmpSal("SH_SALARY") = medsalary.Text
        rsEmpSal("SH_CURRENCYINDI") = clpCode(0).Text
        rsEmpSal("SH_SALCD") = "H"
        rsEmpSal("SH_GRADE") = "00"
        rsEmpSal("SH_EDATE") = dlpStartDate.Text
        rsEmpSal("SH_PAYP") = "T"
        rsEmpSal("SH_SECTION") = clpCode(4).Text
        'rsEmpSal("SH_BAND") = "H"
        rsEmpSal("SH_LDATE") = Date
        rsEmpSal("SH_LTIME") = Time$
        rsEmpSal("SH_LUSER") = glbUserID
        rsEmpSal.Update
        rsEmpSal.Close
        Set rsEmpSal = Nothing
        
        
        'Increase the employee count
        X% = modECount(True)
        
        If Not AUDITDEMO("A") Then MsgBox "ERROR : AUDIT FILE"
        
        glbLEE_SName = txtSurname
        glbLEE_FName = txtFName
    End If
    

    fglbNewEE = False
    glbWFCContractEmployee = False
        
    Unload Me
End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comCountry_Change()
    Call SetCountries
End Sub

'Public Sub cmdClose_Click()
'
'    On Error GoTo err_Unload
'
'    glbTERM_ID = 0
'    glbTran_ID = 0
'    glbTran_Seq = 0
'    glbOnTop = ""
'
'
'    frmEHSINCIDENT.txtDemo(2) = clpCode(2).Text
'    frmEHSINCIDENT.txtDemo(3) = clpCode(3).Text
'    frmEHSINCIDENT.txtDemo(4) = clpCode(4).Text
'    frmEHSINCIDENT.txtDemo(5) = clpCode(5).Text
'    frmEHSINCIDENT.txtDemo(6) = clpCode(6).Text
'    frmEHSINCIDENT.txtDemo(7) = clpCode(7).Text
'    If glbLinamar Then
'        If Mid(frmEHSINCIDENT.txtDemo(8), 4) <> clpCode(8).Text Then
'            frmEHSINCIDENT.txtDemo(8) = clpCode(3).Text & clpCode(8).Text
'        End If
'    Else
'        frmEHSINCIDENT.txtDemo(8) = clpCode(8).Text
'    End If
'
'    If glbLinamar Then
'        If Mid(frmEHSINCIDENT.txtDemo(9), 4) <> clpCode(9).Text Then
'            frmEHSINCIDENT.txtDemo(9) = clpCode(3).Text & clpCode(9).Text
'        End If
'    Else
'        frmEHSINCIDENT.txtDemo(9) = clpCode(9).Text
'    End If
'
'    frmEHSINCIDENT.txtDemo(10) = clpCode(10).Text
'    frmEHSINCIDENT.txtDemo(11) = txtCountryOfEmp
'    If glbLinamar Then
'        If Len(clpHOME(1).Text) > 0 Then
'            frmEHSINCIDENT.txtDemo(12) = clpCode(3).Text & clpHOME(1)
'        End If
'        If Len(clpHOME(2).Text) > 0 Then
'            frmEHSINCIDENT.txtDemo(13) = clpCode(3).Text & clpHOME(2)
'        End If
'    End If
'    frmEHSINCIDENT.txtDemo(14) = txtJobDesc.Text
'    Unload Me
'
'    Exit Sub
'
'err_Unload:
'    Unload Me
'    Resume Next
'    Unload Me
'
'End Sub

Private Sub comCountry_Click()
    Call SetCountries
    txtCountry = comCountry.Text
End Sub

Private Sub comCountry_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comCountry_LostFocus()
    Call SetCountries
End Sub

Private Sub comPayPer_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpEndDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub dlpStartDate_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub elpReptAuthShow_Change(Index As Integer)
txtReptAuthority(Index).Text = getEmpnbr(elpReptAuthShow(Index).Text)
End Sub

Private Sub elpReptAuthShow_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMECONTEMPDEMO"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMECONTEMPDEMO"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim rsTerm As New ADODB.Recordset
Dim X%, SQLQ

glbOnTop = "FRMECONTEMPDEMO"

Screen.MousePointer = HOURGLASS

flagFrmLoad = True

Call setCaption(lblTitle(1))
Call setCaption(lblTitle(14))
lblTitle(20).Caption = lStr(lblTitle(20))

Call setCaption(lblTitle(3))
Call setCaption(lblTitle(4))
Call setCaption(lblTitle(2))
Call setCaption(lblTitle(8))
Call setCaption(lblTitle(9))
Call setCaption(lblIPhone)

lblReptAuth(0).Caption = lStr("Rept. Authority 1")
lblHrsDay.Caption = lStr("Hours/Day")
lblHrsWeek.Caption = lStr("Hours/Week")
lblHrsPayPeriod.Caption = lStr("Hours/Pay Period")

Call addItems

comPayPer.Clear
comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour "
comPayPer.AddItem "Monthly "
comPayPer.AddItem "Daily "

frmEContEmpDemo.Enabled = True

clpJob.TextBoxWidth = 1315
clpJob.Enabled = False

Call INI_Controls(Me)

If Len(txtCountry.Text) > 0 Then comCountry = txtCountry

MDIMain.panHelp(1).Caption = " "

Screen.MousePointer = DEFAULT

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

Private Sub Form_Resize()
    'If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
    '    Me.Left = 0
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call cmdCancel_Click
    Set frmEContEmpDemo = Nothing
    glbOnTop = ""
End Sub

Private Sub UPDMOD()
    'Dim x%
    'For x% = 0 To 2
    '    dlpDate(x%).Enabled = False
    'Next
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
RelateMode = RelateEMP
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
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNewEE Then
    UpdateState = NewRecord
    TF = True
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

'Call ST_UPD_MODE(TF)



'UpdateState = OPENING
'TF = True
'Call set_Buttons(UpdateState)
'If Not UpdateRight Then
'    TF = False
'    Call UPDMOD
'End If

End Sub

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

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
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
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

Private Sub addItems()
Dim ctylist, X

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountry.AddItem Left(ctylist, X - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountry.AddItem ctylist
    End If
Loop

comCountry.ListIndex = 0        '

End Sub

Private Sub medCellPhone_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medHours_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medPCode_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medsalary_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub medTelephone_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtAdd1_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtAdd2_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCity_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCountry_Change()
    Me.comCountry = txtCountry
End Sub

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim dynaJobHIS As New ADODB.Recordset
On Error GoTo JobHis_Err
fglbJobList = ""
Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If
If Not dynaJobHIS.EOF Then
    Do Until dynaJobHIS.EOF
        If Not IsNull(dynaJobHIS!JH_JOB) Then
            fglbJobList = fglbJobList & dynaJobHIS!JH_JOB & ","
        End If
        dynaJobHIS.MoveNext
    Loop
    If Right(fglbJobList, 1) = "," Then
        fglbJobList = Left(fglbJobList, Len(fglbJobList) - 1)
    End If
    dynaJobHIS.MoveFirst
End If
Screen.MousePointer = DEFAULT

Exit Sub

JobHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_JOB_History", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub

Function GetJobDesc(xCode)
Dim SQLQ As String
Dim xDesc As String
Dim dynaJobHIS As New ADODB.Recordset
    xDesc = ""
    If Len(xCode) > 0 Then
        SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
        If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
        dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not dynaJobHIS.EOF Then
            xDesc = dynaJobHIS("JB_DESCR")
        End If
        dynaJobHIS.Close
    End If
    GetJobDesc = xDesc
End Function


Private Function chk_EMPNBR() As Boolean
'''On Error GoTo Eh

    chk_EMPNBR = False
    
    Dim rs As New ADODB.Recordset
    Dim SQLQ As String
    Dim strMsg As String
    Dim retVal As Long
    
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        strMsg = "This Employee Number is already assigned to " & rs("ED_SURNAME") & "," & rs("ED_FNAME") & vbCrLf
        strMsg = strMsg & "Do you want to assign a new number?"
        retVal = MsgBox(strMsg, vbQuestion + vbYesNo, "Duplicate Employee Number")
        If retVal = vbYes Then
            glbLEE_ID = WFC_GenerateContractEmployeeNo
            'If chk_EMPNBR Then
            '    lblEEID = glbLEE_ID
            'End If
        ElseIf retVal = vbNo Then
            Exit Function
        End If
    End If
    rs.Close
    
    chk_EMPNBR = True
    
exH:
    Exit Function
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chk_EMPNBR", "HREMP", "SELECT")
    Resume exH
   
End Function

Private Function WFC_GenerateContractEmployeeNo()
    Dim rsDivContEmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim xNextContNo As Long
        
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDiv.Text & "'"
    rsDivContEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDivContEmp.EOF Then
        If Not IsNull(rsDivContEmp("DV_LSTNUM")) And rsDivContEmp("DV_LSTNUM") <> "" Then
            xNextContNo = rsDivContEmp("DV_LSTNUM")
            
NextContractNo:
            xNextContNo = xNextContNo + 1
                        
            'Check if Employee # already exists
            If WFC_ContractEmpNoExists(xNextContNo) Then
                GoTo NextContractNo
            Else
                rsDivContEmp("DV_LSTNUM") = xNextContNo
            End If
        Else
            'Generate # and # format
            xNextContNo = clpDiv.Text & Format("1", "0000")
            
NextContractNo1:
            'Check if Employee # already exists
            If WFC_ContractEmpNoExists(xNextContNo) Then
                xNextContNo = xNextContNo + 1

                GoTo NextContractNo1
            Else
                rsDivContEmp("DV_LSTNUM") = xNextContNo
            End If
        End If
        rsDivContEmp("LDate") = Date
        rsDivContEmp("LTime") = Time$
        rsDivContEmp("LUser") = glbUserID
        rsDivContEmp.Update
        WFC_GenerateContractEmployeeNo = rsDivContEmp("DV_LSTNUM")
    Else
        WFC_GenerateContractEmployeeNo = ""
    End If
    rsDivContEmp.Close
    Set rsDivContEmp = Nothing
End Function

Private Function WFC_ContractEmpNoExists(xEmpNo) As Boolean
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String

    WFC_ContractEmpNoExists = False
    
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsHREmp.EOF Then
        WFC_ContractEmpNoExists = True
    Else
        WFC_ContractEmpNoExists = False
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
   
End Function

'Private Function WFC_GenerateContractEmployeeNo()
'    Dim rsDivContEmp As New ADODB.Recordset
'    Dim rsHREmp As New ADODB.Recordset
'    Dim SQLQ As String
'
'    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDiv.Text & "'"
'    rsDivContEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsDivContEmp.EOF Then
'        If Not IsNull(rsDivContEmp("DV_LSTNUM")) And rsDivContEmp("DV_LSTNUM") <> "" Then
'            rsDivContEmp("DV_LSTNUM") = rsDivContEmp("DV_LSTNUM") + 1
'        Else
'            rsDivContEmp("DV_LSTNUM") = clpDiv.Text & Format("1", "0000")
'        End If
'        rsDivContEmp("LDate") = Date
'        rsDivContEmp("LTime") = Time$
'        rsDivContEmp("LUser") = glbUserID
'        rsDivContEmp.Update
'        WFC_GenerateContractEmployeeNo = rsDivContEmp("DV_LSTNUM")
'    Else
'        'Add
'        'WFC_GenerateContractEmployeeNo = clpDIv.Text & "9999"
'        WFC_GenerateContractEmployeeNo = ""
'    End If
'    rsDivContEmp.Close
'    Set rsDivContEmp = Nothing
'
'End Function

Private Function WFC_CancelContractEmployeeNo()
    Dim rsDivContEmp As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim SQLQ As String
        
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & clpDiv.Text & "'"
    rsDivContEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDivContEmp.EOF Then
        If Not IsNull(rsDivContEmp("DV_LSTNUM")) And rsDivContEmp("DV_LSTNUM") <> "" Then
            'Make sure no employee with this # already exists just in case.
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHREmp.EOF Then
                'No employee exists with this #
                If glbLEE_ID = rsDivContEmp("DV_LSTNUM") Then
                    rsDivContEmp("DV_LSTNUM") = rsDivContEmp("DV_LSTNUM") - 1
                End If
            End If
            rsHREmp.Close
            Set rsHREmp = Nothing
        End If
        rsDivContEmp("LDate") = Date
        rsDivContEmp("LTime") = Time$
        rsDivContEmp("LUser") = glbUserID
        rsDivContEmp.Update
    End If
    rsDivContEmp.Close
    Set rsDivContEmp = Nothing
End Function

Private Sub SetCountries()
      
If UCase(comCountry) = "CANADA" Then
    lblTitle(18) = "Province"                    '
    clpProv.Tag = "31-Province - Code"          '
    
    lblTitle(19) = "Postal Code"                 '
    medPCode.MaxLength = 7
    medPCode.Mask = "?#? #?#"
    medPCode.Tag = "01-Postal Code"
    
    medTelephone.MaxLength = 14
    medTelephone.Mask = "(###) ###-####"
    
    medCellPhone.MaxLength = 14
    medCellPhone.Mask = "(###) ###-####"
    
ElseIf comCountry = "U.S.A." Then
    
    lblTitle(18) = "State"
    clpProv.Tag = "31-State - Code"         '
    
    lblTitle(19) = "Zip Code"
    medPCode.MaxLength = 10
    medPCode.Mask = "AAAAA-AAAA"
    medPCode.Tag = "01-Zip Code"            '
        
    medTelephone.MaxLength = 14
    medTelephone.Mask = "(###) ###-####"
    
    medCellPhone.MaxLength = 14
    medCellPhone.Mask = "(###) ###-####"

ElseIf comCountry = "MEXICO" Then
    
    lblTitle(18) = "State"
    clpProv.Tag = "31-State - Code"         '
    
    lblTitle(19) = "Zip Code"
    medPCode.MaxLength = 10
    medPCode.Mask = "AAAAA-AAAA"
    medPCode.Tag = "01-Zip Code"            '
        
    If glbLinamar Then
        medTelephone.MaxLength = 25
        medTelephone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
        
        medCellPhone.MaxLength = 25
        medCellPhone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    Else
        medTelephone.MaxLength = 25
        medTelephone.Mask = "(###) ###-####"
        
        medCellPhone.MaxLength = 25
        medCellPhone.Mask = "(###) ###-####"
    End If
    
ElseIf UCase(comCountry) = "BAHAMAS" Then
    lblTitle(18) = "Island"                      '
    clpProv.Tag = "30-Island - Code"            '
    
    lblTitle(19) = "Postal Code"                 '
    medPCode.MaxLength = 8
    medPCode.Mask = "AAAAAAAA"
    medPCode.Tag = "01-Postal Code"             '
    
    medTelephone.MaxLength = 25
    medTelephone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"

    medCellPhone.MaxLength = 25
    medCellPhone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    
Else
    lblTitle(18) = "Province"                '
    clpProv.Tag = "31-Province - Code"      '
    
    lblTitle(19) = "Postal Code"             '
    medPCode.Mask = "&&&&&&&&&&&&&&&"
    medPCode.MaxLength = 15 ' 10
    medPCode.Tag = "01-Postal Code"         '
        
    medTelephone.MaxLength = 25
    medTelephone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
    
    medCellPhone.MaxLength = 25
    medCellPhone.Mask = "&&&&&&&&&&&&&&&&&&&&&&&&&"
End If

End Sub

Private Function NetworkLoginGenerator()
Dim xSurname, xFName
Dim xStr, xLogID
Dim I As Integer

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
    xLogID = getWFCNetworkLoginNoDupicate(xFName, xSurname, "", glbLEE_ID, xStr)
    
    NetworkLoginGenerator = xLogID

End Function

Private Sub txtEmail_GotFocus()
Call SetPanHelp(Me.ActiveControl)
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

Private Sub txtIPHONE_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtReptAuthority_Change(Index As Integer)
elpReptAuthShow(Index) = ShowEmpnbr(txtReptAuthority(Index).Text)
End Sub

Private Sub txtSurname_Change()
If flagFrmLoad = False Then Exit Sub  'carmen may 00
If Len(txtSurname.Text) > 0 Then  ' dont do on add new until in
    Me.lblEEName = RTrim$(txtSurname) & ", " & RTrim$(txtFName)
End If
End Sub

Private Sub txtSurname_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub SetOtherFieldsFromDiv(xDiv)
Dim rsODiv As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ
        
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
    rsODiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsODiv.EOF Then
        If Not IsNull(rsODiv("DV_COUNTRY")) Then
            txtCountry = rsODiv("DV_COUNTRY")
            comCountry = txtCountry
        End If
        If Not IsNull(rsODiv("DV_LOC")) Then
            clpCode(1) = rsODiv("DV_LOC")
        End If
        If Not IsNull(rsODiv("DV_SECTION")) Then
            clpCode(4) = rsODiv("DV_SECTION")
        End If
        If Not IsNull(rsODiv("DV_REGION")) Then
            clpCode(2) = rsODiv("DV_REGION")
        End If
        'If Not IsNull(rsODiv("DV_ADMINBY")) Then
        '    clpCode(3) = rsODiv("DV_ADMINBY")
        'End If
        'If Not IsNull(rsODiv("DV_BONUSDEPT")) Then
        '    If glbWFC Then
        '        'Ticket #27609 Franks 10/07/2015 - comment it out
        '        'If Not glbWFCHrsSal Then
        '        '    txtDeptBonusCtr = rsODiv("DV_BONUSDEPT")
        '        'End If
        '    Else
        '        txtDeptBonusCtr = rsODiv("DV_BONUSDEPT")
        '    End If
        'End If
        'If glbWFC Then 'Ticket #28637 Franks 05/18/2016
        '    If Not IsNull(rsODiv("DV_ORGT1")) Then
        '        clpCode(6).Text = rsODiv("DV_ORGT1")
        '    End If
        'End If
    End If
    rsODiv.Close
    Set rsODiv = Nothing
    
    'Ticket #30359 Franks 07/11/2017 - begin
    'get Position Code based on Div
    SQLQ = "SELECT * FROM HRJOB WHERE LEFT(JB_CODE,7)= '" & xDiv & "IND' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsJOB.EOF Then
        clpJob.Text = rsJOB("JB_CODE")
    End If
    rsJOB.Close
    medHours(0) = 0
    medHours(1) = 0
    medHours(2) = 0
    'Ticket #30359 Franks 07/11/2017 - end
    
End Sub

Private Function getWFCCurrencyIndi(xPlantCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xStr
Dim retVal

    retVal = ""
    If Len(xPlantCode) > 0 Then
        SQLQ = "select * from WFC_Salary_Administration "
        SQLQ = SQLQ & " WHERE SectionCode ='" & xPlantCode & "' "
        SQLQ = SQLQ & " AND NOT ( CurrencyIndicator IS NULL OR CurrencyIndicator = '') "
        SQLQ = SQLQ & "ORDER BY FiscalYear DESC"
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp("CurrencyIndicator")) Then
                retVal = rsTemp("CurrencyIndicator")
            End If
        End If
        rsTemp.Close
        Set rsTemp = Nothing
    End If
    getWFCCurrencyIndi = retVal
End Function
