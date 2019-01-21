VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEPERFORM 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Performance"
   ClientHeight    =   10950
   ClientLeft      =   405
   ClientTop       =   1365
   ClientWidth     =   12090
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
   ScaleHeight     =   10950
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel panWindow 
      Height          =   5295
      Left            =   0
      TabIndex        =   27
      Top             =   2040
      Width           =   8895
      _Version        =   65536
      _ExtentX        =   15690
      _ExtentY        =   9340
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
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   120
         ScaleHeight     =   5535
         ScaleWidth      =   8895
         TabIndex        =   28
         Top             =   120
         Width           =   8895
         Begin VB.CommandButton cmdImport 
            Caption         =   "Import"
            Height          =   270
            Left            =   7830
            TabIndex        =   11
            Top             =   2460
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtReptAuthority 
            Appearance      =   0  'Flat
            DataField       =   "PH_REPTAU"
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
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   12
            Tag             =   "00-Employee Number of individual's supervisor"
            Top             =   1410
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox txtReptAuthority 
            Appearance      =   0  'Flat
            DataField       =   "PH_REPTAU2"
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
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   13
            Tag             =   "00-Employee Number of individual's supervisor"
            Top             =   1740
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox txtReptAuthority 
            Appearance      =   0  'Flat
            DataField       =   "PH_REPTAU3"
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
            Index           =   2
            Left            =   2010
            MaxLength       =   12
            TabIndex        =   14
            Tag             =   "00-Employee Number of individual's supervisor"
            Top             =   2070
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox memComments 
            Appearance      =   0  'Flat
            DataField       =   "PH_COMMENTS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1350
            Left            =   2010
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Tag             =   "00-Enter Comments"
            Top             =   3060
            Width           =   6795
         End
         Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   6
            Tag             =   "00-Employee Number of individual's supervisor"
            Top             =   2070
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   503
            ShowUnassigned  =   1
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   5
            Tag             =   "00-Employee Number of individual's supervisor"
            Top             =   1740
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   503
            ShowUnassigned  =   1
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   4
            Tag             =   "00-Employee Number of individual's supervisor"
            Top             =   1410
            Width           =   4995
            _ExtentX        =   8811
            _ExtentY        =   503
            ShowUnassigned  =   1
            RefreshDescriptionWhen=   2
         End
         Begin INFOHR_Controls.CodeLookup clpPosCode 
            DataField       =   "PH_JOB"
            Height          =   285
            Left            =   1680
            TabIndex        =   0
            Tag             =   "01-Position code"
            Top             =   60
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   5
         End
         Begin INFOHR_Controls.DateLookup dlpReviewDate 
            DataField       =   "PH_PNEXT"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   7
            Tag             =   "40-Next Date to Review Performance"
            Top             =   2400
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpReviewDate 
            DataField       =   "PH_PREVIEW"
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Tag             =   "41-Performance Review Date"
            Top             =   390
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "PH_PCODE"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   2
            Tag             =   "00-Performance Rating - Code "
            Top             =   720
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDPC"
         End
         Begin Threed.SSFrame fraCurrentSalary 
            Height          =   1170
            Left            =   6750
            TabIndex        =   29
            Top             =   480
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   2064
            _StockProps     =   14
            Caption         =   "Current Salary"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Begin VB.Label lblEDateD 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
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
               Height          =   195
               Left            =   930
               TabIndex        =   33
               Top             =   705
               Width           =   840
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Effective "
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
               Left            =   195
               TabIndex        =   32
               Top             =   705
               Width           =   840
            End
            Begin VB.Label lblCSalary 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0.00"
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
               Left            =   930
               TabIndex        =   31
               Top             =   390
               Width           =   315
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Salary"
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
               Left            =   195
               TabIndex        =   30
               Top             =   390
               Width           =   435
            End
         End
         Begin Threed.SSCheck chkCurrent 
            DataField       =   "PH_CURRENT"
            Height          =   255
            Left            =   7230
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   0
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Current Record"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin MSMask.MaskEdBox medSalary 
            DataField       =   "PH_BONUS"
            Height          =   285
            Left            =   1995
            TabIndex        =   8
            Tag             =   "20-Enter bonus $"
            Top             =   2730
            Width           =   1530
            _ExtentX        =   2699
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
            Format          =   "#,##0.00;(#,##0.00)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "PH_PCODE2"
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   3
            Tag             =   "00-Performance Rating - Code "
            Top             =   1065
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDPC"
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "PH_LDATE"
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   10
            Tag             =   "40-Transaction Date"
            Top             =   4560
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            TextBoxWidth    =   1215
            Enabled         =   0   'False
         End
         Begin VB.Label lblUpdateBy 
            Caption         =   "Updated By"
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
            Left            =   0
            TabIndex        =   53
            Top             =   4920
            Width           =   1095
         End
         Begin VB.Label lblUserDesc 
            Caption         =   "lblUserDesc"
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
            Left            =   2040
            TabIndex        =   52
            Top             =   4920
            Width           =   2775
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Date"
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
            Left            =   0
            TabIndex        =   51
            Top             =   4605
            Width           =   1455
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Left            =   7410
            Picture         =   "feperfrm.frx":0000
            Top             =   2460
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Left            =   7410
            Picture         =   "feperfrm.frx":014A
            Top             =   2460
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblImport 
            Alignment       =   1  'Right Justify
            Caption         =   "Performance Review"
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
            Height          =   255
            Left            =   5190
            TabIndex        =   50
            Top             =   2460
            Visible         =   0   'False
            Width           =   1995
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Performance Rating 2"
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
            Left            =   30
            TabIndex        =   49
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bonus $"
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
            Left            =   0
            TabIndex        =   48
            Top             =   2760
            Width           =   585
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority 3"
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
            Left            =   0
            TabIndex        =   47
            Top             =   2100
            Width           =   1650
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority 2"
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
            Left            =   30
            TabIndex        =   46
            Top             =   1770
            Width           =   1650
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Next Review Date"
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
            Left            =   0
            TabIndex        =   45
            Top             =   2430
            Width           =   1560
         End
         Begin VB.Label lblJobID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "JobID"
            DataField       =   "PH_JOB_ID"
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
            Left            =   5670
            TabIndex        =   44
            Top             =   90
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblJob 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2250
            TabIndex        =   43
            Top             =   60
            UseMnemonic     =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   195
            Index           =   6
            Left            =   0
            TabIndex        =   42
            Top             =   3090
            Width           =   870
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reporting Authority 1"
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
            Left            =   15
            TabIndex        =   41
            Top             =   1440
            Width           =   1650
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Performance Rating"
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
            Left            =   15
            TabIndex        =   40
            Top             =   750
            Width           =   1680
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Review Date"
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
            Left            =   15
            TabIndex        =   39
            Top             =   390
            Width           =   1110
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   38
            Top             =   60
            Width           =   690
         End
         Begin VB.Label lblPerfID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "lblPID"
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
            Height          =   225
            Left            =   5520
            TabIndex        =   37
            Top             =   3240
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label lblEEID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "EEId"
            DataField       =   "PH_EMPNBR"
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
            Left            =   6930
            TabIndex        =   36
            Top             =   3585
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblCNum 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comp"
            DataField       =   "PH_COMPNO"
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
            Left            =   6840
            TabIndex        =   35
            Top             =   2850
            Visible         =   0   'False
            Width           =   405
         End
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   5295
      LargeChange     =   315
      Left            =   8880
      Max             =   4575
      SmallChange     =   315
      TabIndex        =   26
      Top             =   2040
      Width           =   255
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   12090
      _Version        =   65536
      _ExtentX        =   21325
      _ExtentY        =   979
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
         Left            =   7200
         TabIndex        =   54
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   22
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
         Left            =   1530
         TabIndex        =   21
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   3120
         TabIndex        =   20
         Top             =   135
         Width           =   1740
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feperfrm.frx":0294
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "feperfrm.frx":02A8
      TabIndex        =   15
      Top             =   600
      Width           =   9135
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "PH_LDATE"
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
      Left            =   9240
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "PH_LTIME"
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
      Left            =   9240
      MaxLength       =   25
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "PH_LUSER"
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
      Left            =   9240
      MaxLength       =   25
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   23
      Top             =   10290
      Width           =   12090
      _Version        =   65536
      _ExtentX        =   21325
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
      Begin VB.CommandButton cmdPosition 
         Appearance      =   0  'Flat
         Caption         =   "P&osition"
         Height          =   280
         Left            =   10320
         TabIndex        =   25
         Top             =   30
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.CommandButton cmdSalary 
         Appearance      =   0  'Flat
         Caption         =   "&Salary"
         Height          =   280
         Left            =   10320
         TabIndex        =   24
         Top             =   330
         Visible         =   0   'False
         Width           =   1250
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9660
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8610
         Top             =   30
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   1
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
Attribute VB_Name = "frmEPERFORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsnapEENames As New ADODB.Recordset
Dim Job_Snaps As New ADODB.Recordset
Dim dynaJobHIS As New ADODB.Recordset
Dim dynaSalHIS As New ADODB.Recordset

Dim fglbJobDesc$
Dim fglbJobID&
Dim savAuth(3)

Dim fglbCurSalary@, fglbCurSalEdate  As Variant
Dim fglbNew As Integer    '
Dim orgReviewDate As String 'Ticket #21601 Franks 02/24/2012
Dim orgNextReviewDate As String  'added by Laura 11/2/97
Dim orgComments As String '
Dim orgPosCode As String
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbJobList
Dim flgloaded As Boolean
Dim locCountry As String
Dim MailBody

Private Function chkPerformance()
Dim SQLQ As String, Msg$, dd&, Response%
Dim DgDef As Variant, Title$, DCurPDate As Variant
Dim X% 'Jaddy 10/28/99
chkPerformance = False

On Error GoTo chkPerH_Err

If Len(dlpReviewDate(0).Text) = 0 And Len(dlpReviewDate(1).Text) = 0 And Len(clpCode(1).Text) = 0 Then
    Msg$ = "Must enter one of Review Date or"
    Msg$ = Msg$ & Chr(10) & "Next Review Date or " & lStr("Performance") & " rating"
    DgDef = MB_OKCANCEL '+ MB_ICONQUESTION + MB_DEFBUTTON2
    Response% = MsgBox(Msg$, DgDef, "Warning!")
        If Response% = IDOK Then
            dlpReviewDate(0).SetFocus
            Exit Function
        Else
            Unload Me
            Exit Function
        End If
End If

If Len(dlpReviewDate(0).Text) > 0 Then
    If Not IsDate(dlpReviewDate(0).Text) Then
        Msg$ = "Not a valid Review Date"
        dlpReviewDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    Else
        If glbSetPer Then
            DCurPDate = CurPDate()
            If DCurPDate > 0 Then    ' 0 if no current record out there
                DCurPDate = CVDate(DCurPDate)
                If DateDiff("d", CVDate(dlpReviewDate(0).Text), DCurPDate) <= 0 Then
                    Msg$ = "Warning...you cannot add or edit a record with a date"
                    Msg$ = Msg$ & Chr(10) & "the same or later than your most current record."
                    Msg$ = Msg$ & Chr(10) & "If you need to edit current " & lStr("Performance") & ", "
                    Msg$ = Msg$ & Chr(10) & "go to " & lStr("Performance") & " screen under Employee Menu."
                    MsgBox Msg$
                    dlpReviewDate(0).SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If

If glbLinamar Then   'Linamar - Ticket #15746
    If Len(dlpReviewDate(1).Text) = 0 Then
        Msg$ = "Next Review Date is required"
        dlpReviewDate(1).SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If

If Len(dlpReviewDate(1).Text) > 0 Then
    If Not IsDate(dlpReviewDate(1).Text) Then
        Msg$ = "Not a valid Next Review Date"
        dlpReviewDate(1).SetFocus
        MsgBox Msg$
        Exit Function
    Else
        If IsDate(dlpReviewDate(0).Text) Then  'jdy 5/2/00
            dd& = DateDiff("d", CVDate(dlpReviewDate(0).Text), CVDate(dlpReviewDate(1).Text))
            If dd& < 0 Then
                Msg$ = "Next Review date can not preceed this Review Date."
                dlpReviewDate(1).SetFocus
                MsgBox Msg$
                Exit Function
            End If
        End If
    End If
End If

If fglbNew = True And (Not glbSetPer) Then    'laura dec 08, 1997; added just this if
    If glbAddHisWarning Then
        DCurPDate = CurPDate()
        If DCurPDate > 0 Then    ' 0 if no current record out there
            DCurPDate = CVDate(DCurPDate)
            If Len(dlpReviewDate(0).Text) > 0 Then
                If DateDiff("d", CVDate(dlpReviewDate(0).Text), DCurPDate) >= 0 Then
                    Msg$ = "Warning, you can not add a record with a date"
                    Msg$ = Msg$ & Chr(10) & "the same or earlier than your most current record."
                    'Msg$ = Msg & Chr(10) & "Do you want to proceed?"
                    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
                    Response% = MsgBox(Msg$) ', DgDef, "Warning!")
                    'If Response% = IDNO Then
                    dlpReviewDate(0).SetFocus
                    Exit Function
                    'End If
                End If
            End If
        End If
    End If
End If


If Len(clpCode(1).Text) > 0 Then
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "If Code Entered Must Be Valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If glbWFC Then 'Ticket #29633 Franks 01/06/2016
    If clpCode(1).Text = "01" Or clpCode(1).Text = "02" Or clpCode(1).Text = "03" Then
        MsgBox "You can not select '01', '02' or '03' as Performance Rating code, they are Incentive Scorecard Codes"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

chkPerformance = True

Exit Function

chkPerH_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkPerf", "HR_PERFORM_HISTORY", "edit/Add")
Call RollBack '28July99 js

End Function

Private Function Chkpos() 'Laura nov 24, 1997
Dim SQLQ As String, Msg$, X%
Dim Snap_Job_His As New ADODB.Recordset


On Error GoTo ChkPosPerf_Err

Chkpos = False

If Len(clpPosCode.Text) > 0 Then
    If clpPosCode.Caption = "Unassigned" Then
        MsgBox "Position Code is invalid"
         clpPosCode.SetFocus
        Exit Function
    End If
Else
    If clpPosCode.Caption = "Unassigned" Then
        MsgBox "Position Code is required"
         clpPosCode.SetFocus
        Exit Function
    End If
End If
If Not Set_Position(clpPosCode.Text, False) Then
    Msg$ = "No position found "
    Msg$ = Msg$ & Chr(10) & "Please review positions from Job History!"
    MsgBox Msg$
    Exit Function
Else
For X% = 0 To 2
    If elpReptAuthShow(X%).Text = "0" Then elpReptAuthShow(X%).Text = ""
    If elpReptAuthShow(X%).Enabled Then
        If Len(elpReptAuthShow(X%).Text) > 0 Then
            If elpReptAuthShow(X%).Caption = "Unassigned" Then
                MsgBox "Employee # not valid. Check # and re-enter!"
                If elpReptAuthShow(X%).Enabled Then elpReptAuthShow(X%).SetFocus
                Exit Function
            End If
        End If
    End If
Next
End If

If glbCompSerial = "S/N - 2259W" Then 'Oxford 'Ticket #21599 Franks 03/01/2012
    If glbMulti Then
        If fglbNew Then 'new record
            If CheckDuplCurrent(glbLEE_ID, clpPosCode.Text) Then
                Msg$ = "There is another current Performance for the same Position Code '" & clpPosCode.Text & "' " & Chr(10)
                Msg$ = Msg$ & "You can't have two current Performances for the same Position Code" & Chr(10)
                Msg$ = Msg$ & "Please uncheck the Current Record flag for the previous Current Performance." & Chr(10)
                MsgBox Msg$
                Exit Function
            End If
        End If
    End If
End If

Chkpos = True
Exit Function

ChkPosPerf_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdPosCode", "HR_JOB_HISTORY", "Change Position")
Resume Next

End Function

Sub cmdCancel_Click()
Dim X As Integer
On Error GoTo Can_Err
fglbNew = False

rsDATA.CancelUpdate
Call Display_Value
Call SET_UP_MODE

For X = 0 To 2
    Call txtReptAuthority_Change(X)
Next
'Call ST_UPD_MODE(True)  ' reset screen's attributes

'fglbNew = False
'Call SET_UP_MODE

'Ticket #24099
If glbCompSerial <> "S/N - 2368W" Then
    'Ticket #24099 - Show the Current Salary and Effective Date
    'If Set_Salary(clpPosCode, True) Then
    If Set_Salary("", True) Then
        lblCSalary = Round2DEC(fglbCurSalary@)
        lblEDateD = fglbCurSalEdate
    Else
        lblCSalary = ""
        lblEDateD = "Unassigned"
    End If
Else
    lblCSalary = ""
    lblEDateD = ""
End If

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_PERFORM_HISTORY", "Cancel")
Call RollBack '28July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEPERFORM" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

orgReviewDate = dlpReviewDate(0).Text
orgNextReviewDate = dlpReviewDate(1).Text
If orgNextReviewDate <> "" Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If

glbJob = rsDATA("PH_JOB")
glbSDate = rsDATA("PH_PREVIEW")

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
'George Jan 26,2006
If gsAttachment_DB Then
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute "Delete from HRDOC_PERFORM_HISTORY where DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR = " & glbLEE_ID & " and DH_JOB='" & glbJob & "' and DH_PREVDATE=" & Date_SQL(glbSDate)
    gdbAdoIhr001_DOC.CommitTrans
End If
'George Jan 26,2006
Data1.Refresh

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Call Set_Current_Flag
End If
Call Display_Value
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_PERFORM_HISTORY", "Delete")
Call RollBack '28July99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Response%, Msg$, Title$, DgDef As Double


On Error GoTo Mod_Err

orgReviewDate = dlpReviewDate(0).Text
orgNextReviewDate = dlpReviewDate(1).Text
orgComments = memComments
orgPosCode = clpPosCode.Text

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_PERFORM_HISTORY", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String, Msg$
Dim X
On Error GoTo AddN_Err

If Not Set_Position("", True) Then
    Msg$ = "No current position found "
    Msg$ = Msg$ & Chr(10) & "Please review position prior to updating salary."
    MsgBox Msg$
    Exit Sub
End If
If glbCompSerial <> "S/N - 2368W" Then
    If Not Set_Salary("", True) Then
        Msg$ = "No current salary record found "
        Msg$ = Msg$ & Chr(10) & "Please review salary history."
        glbStopPerform% = True
        MsgBox Msg$
        Exit Sub
    End If
Else
    fglbCurSalary = 0
    fglbCurSalEdate = ""
End If
If Not getJOB() Then
    Msg$ = "Can not find description for current position."
    Msg$ = Msg$ & Chr(10) & "Please review position Information list."
    MsgBox Msg$
    Exit Sub    ' no damage yet
End If

If glbMulti Then
    Call CR_JobHis_Snap
    Call CR_SalHis_Snap
    'fglbJobList = ""
    clpPosCode.seleEMPCode = fglbJobList
End If

If glbCompSerial <> "S/N - 2368W" Then
    lblCSalary = Round2DEC(fglbCurSalary@)
    lblEDateD = fglbCurSalEdate
Else
    lblCSalary = ""
    lblEDateD = ""
End If
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

'George on Jan 26,2006 #10266
If gsAttachment_DB Then
    glbJob = ""
    glbSDate = "01/01/1900"
    lblImport.Visible = True
    imgSec.Visible = False
    imgNoSec.Visible = True
    cmdImport.Visible = True
End If
'George on Jan 26,2006 #10266

'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)
rsDATA.AddNew

clpPosCode.Text = lblJob
orgPosCode = clpPosCode.Text
chkCurrent = glbMulti
lblJobID = fglbJobID&
clpCode(1).Text = fglbJobDesc$
lblEEID = glbLEE_ID
lblCNum.Caption = "001"
For X = 0 To 2
    elpReptAuthShow(X).Text = ShowEmpnbr(savAuth(X))
Next
dlpReviewDate(1).Text = ""
dlpReviewDate(0).SetFocus
dlpDate(1).Text = Format(Now, "SHORT DATE")
Updstats(0).Text = Format(Now, "SHORT DATE")
'If glbSetPer Or glbMulti Then clpPosCode.SetFocus

If glbWFC Then 'Ticket #29011 Franks 08/03/2016 - begin
    clpCode(2).Text = "3"
End If

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PERFORM_HISTORY", "Add")
Call RollBack '28July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim rsPER As New ADODB.Recordset
Dim X, xReptAuthority, xID
Dim xBranch

On Error GoTo Add_Err

'City of Timmins - Ticket #13207
If glbCompSerial = "S/N - 2375W" And fglbNew <> True Then
    'Ask for the password
    glbAccessPswd = False
    frmAccessPswd.Show 1
    If glbAccessPswd = False Then   'Access Denied
        Call cmdCancel_Click
        Exit Sub
    End If
End If

'Franks May 8,2003 Cause error of 'Row cannot be located for updateing'
'Move chkPerformance to top, before it was behind clpPostCode
If Not chkPerformance() Then Exit Sub

If clpPosCode.Enabled Then
    If Not Chkpos() Then Exit Sub
End If

'If Not chkPerformance() Then Exit Sub
Screen.MousePointer = HOURGLASS

'Call UpdUStats(Me) ' update user's stats (who did it and when)
Updstats(1).Text = Time$
Updstats(2).Text = glbUserID

Call Set_Control("U", Me, rsDATA)

For X = 0 To 2
    xReptAuthority = getEmpnbr(elpReptAuthShow(X).Text)
    rsDATA("PH_REPTAU" & IIf(X = 0, "", X + 1)) = IIf(Val(xReptAuthority) = 0, Null, xReptAuthority)
Next

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA!PH_ID
    glbJob = rsDATA("PH_JOB")
    glbSDate = rsDATA("PH_PREVIEW")
    rsDATA.Requery
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_PERFORM_HISTORY set DH_PREVDATE=" & Date_SQL(rsDATA("PH_PREVIEW")) & " where DH_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DH_JOB='" & glbJob & "' and DH_PREVDATE=" & Date_SQL(glbSDate) & " AND DH_DOCKEY = " & glbDocKey
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA!PH_ID
    glbJob = rsDATA("PH_JOB")
    glbSDate = rsDATA("PH_PREVIEW")
    rsDATA.Requery
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update HRDOC_PERFORM_HISTORY set DH_PREVDATE=" & Date_SQL(rsDATA("PH_PREVIEW")) & " where DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR = " & glbLEE_ID & " and DH_JOB='" & glbJob & "' and DH_PREVDATE=" & Date_SQL(glbSDate) & " AND DH_DOCKEY = " & glbDocKey
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
End If

'Data1.Refresh

Call Set_Current_Flag

'Line_chkCurrent_1
'If chkCurrent Then
'    If Not updFollow("U") Then Exit Sub
'End If

Data1.Refresh
Data1.Recordset.Find "PH_ID=" & xID

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
        End If
    End If
    glbDocImpFile = ""
End If

Call Display_Value

'Ticket #20190 Franks 04/20/2011
If glbCompSerial = "S/N - 2355W" Then 'County of Lambton
    If glbAdv Then
        Call Employee_Master_Integration(glbLEE_ID)
    End If
End If
'If gsEMAIL_ONPERFORMANCE Then
'Ticket #22409 Frank 08/08/2012 - add "not glbWFC" since they use this email for Smoker Status Change
If gsEMAIL_ONPERFORMANCE And Not glbWFC Then
    MailBody = ""
    If NewHireForms.count = 0 Then 'Non new hire
        'If fglbNew Or chkCurrent Then
        If isReviewDatesChanged Then
            MailBody = "The " & lStr("Performance") & " has been changed." & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
                If Len(xBranch) > 0 Then
                    xBranch = GetTABLDesc("EDSE", xBranch)
                    MailBody = MailBody & "Branch: " & xBranch & vbCrLf
                End If
            End If
            If IsDate(dlpReviewDate(0).Text) Then
                MailBody = MailBody & "Review Date: " & dlpReviewDate(0) & vbCrLf
            End If
            If IsDate(dlpReviewDate(1).Text) Then
                MailBody = MailBody & "Next Review Date: " & dlpReviewDate(1) & vbCrLf
            End If
        End If
    End If
End If

'Ticket #22682: Release 8.0 - Set older Performance Review Follow Up records as Completed first if uncompleted
'follow up records are found for Salary, before adding a new follow up record.
If fglbNew And NewHireForms.count = 0 Then
    glbFollowUpList = "PREV"
    If Older_FollowUp_Records_Found(glbFollowUpList) Then
        frmFollowUpList.Show 1
    End If
End If

'Ticket #13079 'Moved the following function from Line_chkCurrent_1 to Line_chkCurrent_2
'Line_chkCurrent_2
If chkCurrent Then
    If Not updFollow("U") Then Exit Sub
End If

fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

'If gsEMAIL_ONPERFORMANCE Then
'Ticket #22409 Frank 08/08/2012 - add "not glbWFC" since they use this email for Smoker Status Change
If gsEMAIL_ONPERFORMANCE And Not glbWFC Then
    If Len(MailBody) > 0 Then
        Screen.MousePointer = DEFAULT
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #21601 Franks 02/24/2012
            Call EmailSendingForSamuel
        Else
            Call imgEmail_Click
        End If
    End If
End If

Screen.MousePointer = DEFAULT
Call NextForm

Exit Sub

Add_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PERFORM_HISTORY", "Update")
Call RollBack '28July99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub



Private Sub cmdPosCode_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPosition_Click()
'Unload frmEPOSITION
'glbSetPos = glbSetPer
'frmEPOSITION.Show
'Unload Me

'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, dscGroup$

'cmdPrint.Enabled = False
RHeading = lblEEName.Caption & "'s " & lStr("Performance") & " History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"

If Not glbtermopen Then
    xReport = glbIHRREPORTS & "rgridper.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    Me.vbxCrystal.SelectionFormula = "{HR_PERFORM_HISTORY.PH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    xReport = glbIHRREPORTS & "rgridpe1.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    End If
        Me.vbxCrystal.SelectionFormula = "{Term_PERFORM_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
End If

Me.vbxCrystal.Formulas(0) = "lblRptTitle = '" & UCase(lStr("Performance")) & " INFORMATION :'"

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, dscGroup$

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False
RHeading = lblEEName.Caption & "'s " & lStr("Performance") & " History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"


If Not glbtermopen Then
    xReport = glbIHRREPORTS & "rgridper.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    Me.vbxCrystal.SelectionFormula = "{HR_PERFORM_HISTORY.PH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    xReport = glbIHRREPORTS & "rgridpe1.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    End If
        Me.vbxCrystal.SelectionFormula = "{Term_PERFORM_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
End If

Me.vbxCrystal.Formulas(0) = "lblRptTitle = '" & UCase(lStr("Performance")) & " INFORMATION :'"


Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub


'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdSalary_Click()
'Unload frmESALARY
'glbSetSal = glbSetPer
'frmESALARY.Show
'Unload Me
'End Sub

Private Sub CR_Job_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo Job_Err

Screen.MousePointer = HOURGLASS

'SQLQ = "SELECT * FROM HRJOB"
SQLQ = "SELECT TOP 10 * FROM HRJOB"  'Ticket #27983 Franks 02/10/2016

If Job_Snaps.State <> 0 Then Job_Snaps.Close
Job_Snaps.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Job_Snaps.EOF And Job_Snaps.BOF Then
    Msg = "No Job descriptions found" & Chr(10)
    Msg = Msg & "You will require authority to add one to continue"
    MsgBox Msg
Else
    'EOF?
    Job_Snaps.MoveFirst
End If

Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Jobs", "HRJOB", "SELECT")
Call RollBack '28July99 js
 
End Sub

Private Function CurPDate()
Dim SQLQ As String
Dim HRP_Snap As New ADODB.Recordset

CurPDate = 0    ' returns 0 if no found records

On Error GoTo JP_Err

SQLQ = "Select HR_PERFORM_HISTORY.* from HR_PERFORM_HISTORY"
SQLQ = SQLQ & " where HR_PERFORM_HISTORY.PH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND HR_PERFORM_HISTORY.PH_CURRENT <>0"

HRP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

If HRP_Snap.BOF And HRP_Snap.EOF Then
    Exit Function
Else
    CurPDate = HRP_Snap("PH_PREVIEW")
    HRP_Snap.Close
End If

Exit Function

JP_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Perform History Snap", "HR_PERFORM_HISTORY", "SELECT")
Call RollBack '28July99 js

End Function

Function EERetrieve()
Dim SQLQ As String
Dim X, xFld
Dim rs As New ADODB.Recordset

EERetrieve = False

On Error GoTo EERError

    Screen.MousePointer = HOURGLASS
    
    'Ticket #28635 - Add View Own security
    If Not glbtermopen Then
        If glbUserEmpNo = glbLEE_ID And Not gSec_Performance_ViewOwn Then
            MsgBox "You cannot view your own " & lStr("Performance") & " information.", vbCritical, "info:HR - Security"
            'glbLEE_ID = 0      'Ticket #25208
            Screen.MousePointer = DEFAULT
            Unload Me: Exit Function
        End If
    End If
    
    If glbCompSerial = "S/N - 2259W" Then 'Added by Bryan 11/07/05 Ticket #8857
        If glbtermopen Then
            SQLQ = "Select ED_SECTION FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        If rs("ED_SECTION") = "Y" Then
            glbMulti = True
        Else
            glbMulti = False
        End If
        rs.Close
        Set rs = Nothing
        SQLQ = ""
    End If
    
    'WDGPHU - Ticket #27899
    If glbCompSerial = "S/N - 2411W" Then
        If glbtermopen Then
            SQLQ = "Select ED_ORGT1 FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_ORGT1 FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        If rs("ED_ORGT1") = "YES" Then
            glbMulti = True
        Else
            glbMulti = False
        End If
        rs.Close
        Set rs = Nothing
        SQLQ = ""
    End If
    
    If glbWFC Then
        If glbtermopen Then
            SQLQ = "Select ED_SECTION,ED_COUNTRY FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_SECTION,ED_COUNTRY FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        Dim rs2 As New ADODB.Recordset
        rs2.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        locCountry = ""
        If Not rs2.EOF Then
            locCountry = rs2("ED_COUNTRY")
        End If
        rs2.Close
        Set rs2 = Nothing
        SQLQ = ""
    End If
    
    If glbtermopen Then
        SQLQ = "SELECT Term_PERFORM_HISTORY.*,"
    Else
        SQLQ = "SELECT HR_PERFORM_HISTORY.*,"
    End If

For X = 0 To 2
    xFld = "REPTAU" & IIf(X = 0, "", X + 1)
    If glbLinamar Then
        SQLQ = SQLQ & " CASE WHEN PH_" & xFld & " IS NOT NULL AND LEN(PH_" & xFld & ")>2 "
        SQLQ = SQLQ & " THEN RIGHT(PH_" & xFld & ",3)+'-'+"
        SQLQ = SQLQ & " LEFT(PH_" & xFld & ",LEN(PH_" & xFld & ")-3) "
        SQLQ = SQLQ & " ELSE STR(PH_" & xFld & ") END "
        SQLQ = SQLQ & " AS " & xFld & IIf(X = 2, "", ",")
    Else
        If glbOracle Then
            SQLQ = SQLQ & "PH_" & xFld & " AS " & xFld & IIf(X = 2, "", ",")
        Else
            SQLQ = SQLQ & "STR(PH_" & xFld & ") AS " & xFld & IIf(X = 2, "", ",")
        End If
        
    End If
Next
If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_PERFORM_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
Else
    SQLQ = SQLQ & " FROM  HR_PERFORM_HISTORY"
    SQLQ = SQLQ & " WHERE PH_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC, PH_PNEXT DESC"
    
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True  'new

Screen.MousePointer = DEFAULT
Call Display_Value
Call SET_UP_MODE

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, lStr("Performance"), "HR_PERFORM_HISTORY", "SELECT")
Call RollBack '28July99 js

Exit Function

End Function

Private Sub clpCode_LostFocus(Index As Integer)
    'Ticket #24564 - Macaulay - Compute the Next Review Date if blank
    If glbCompSerial = "S/N - 2420W" And Not IsDate(dlpReviewDate(1)) Then
        Call NextReviewDate_Macaulay
    End If
End Sub

Private Sub dlpDate_Change(Index As Integer)
    Updstats(0).Text = dlpDate(1).Text
End Sub

Private Sub elpReptAuthShow_Change(Index As Integer)
txtReptAuthority(Index).Text = getEmpnbr(elpReptAuthShow(Index).Text)
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMEPERFORM"
clpPosCode.seleEMPCode = fglbJobList
flgloaded = True
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEPERFORM"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%

glbOnTop = "FRMEPERFORM"

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT


Call setCaption(lblTitle(11))

If glbCompSerial = "S/N - 2387W" Then  'Bird Packaging Limited - Ticket #14060
    lblTitle(12).Caption = "Performance Type"
End If

If glbLinamar Then  'Linamar - Ticket #15746
    lblTitle(5).FontBold = True
End If

If glbWFC Then 'Ticket #17823
    'clpPosCode.TextBoxWidth = 1215 'Ticket #25911 Franks 11/10/2014
    lblTitle(11).Caption = "Lump Sum $"
    vbxTrueGrid.Columns(9).Caption = "Lump Sum $"
    medSalary.Tag = "20-Enter Lump Sum $"
End If
clpPosCode.TextBoxWidth = 1315 'Ticket #26726 Franks 06/15/2015 for all

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    If glbNoNONE Then
        If glbUNION = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then        'Hemu -EXE
        If glbUNION = "EXEC" Then   'Hemu -EXE
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    If glbNoNONE Then
        If glbUNIONTe = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then        'Hemu -EXE
        If glbUNIONTe = "EXEC" Then     'Hemu -EXE
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
End If

'Ticket #28635 - Add View Own Security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_Performance_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Performance") & " information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Sub
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If
Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) < 1 Then Exit Sub
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = IIf(glbSetPer, "Set ", "") & lStr("Performance") & " History - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call CR_JobHis_Snap
Call CR_SalHis_Snap
If Data1.Recordset.EOF Then
    If Not Set_Position("", True) Then Exit Sub
    If glbCompSerial <> "S/N - 2368W" Then
        If Set_Salary("", True) Then
            lblCSalary = Round2DEC(fglbCurSalary@)
            If Not IsNull(fglbCurSalEdate) Then
                lblEDateD = fglbCurSalEdate
            End If
        End If
    Else
        lblCSalary = ""
        lblEDateD = ""
    End If
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call Display_Value
Call SET_UP_MODE
If Not gSec_Upd_Performance Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If
If glbCompSerial = "S/N - 2368W" Then
    fraCurrentSalary.Visible = False
Else
    If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
        'Check if the user has access to this employee's salary information
        If Allow_User_To_View("ACTIVE") = False Then
            fraCurrentSalary.Visible = False
        Else
            fraCurrentSalary.Visible = gSec_Inq_Salary
        End If
    Else
        fraCurrentSalary.Visible = gSec_Inq_Salary
    End If
End If

Call INI_Controls(Me)

'Performance labels
lblTitle(2).Caption = lStr("Performance Rating") 'lStr(lblTitle(2).Caption)
If glbWFC Then
    'Ticket #29011 Franks 08/03/2016 - begin
    lblTitle(12).Caption = "Incentive Scorecard"
    vbxTrueGrid.Columns(2).Caption = "Incentive Scorecard"
    'clpCode(2).TransDiv = "'01','02','03','1','2','3'"
    clpCode(2).TransDiv = "'01','02','03'" 'Ticket #29633 Franks 01/06/2017
    clpCode(2).TABLTitle = UCase("Incentive Scorecard Codes")
    'Ticket #29011 Franks 08/03/2016 - end
Else
    lblTitle(12).Caption = lStr("Performance Rating 2") 'lStr(lblTitle(12).Caption)
End If
If glbCompSerial = "S/N - 2172W" Then 'Ticket #17336 County of Lanark
    lblTitle(2).Caption = ("Performance Type")
    lblTitle(12).Caption = ("Performance Rating")
End If
lblImport.Caption = lStr(lblImport.Caption)

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

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    panWindow.Height = Me.ScaleHeight - (panEEDESC.Height + vbxTrueGrid.Height + 200)
    panWindow.Width = vbxTrueGrid.Width - scrControl.Width
    If Me.Height >= panEEDESC.Height + panDetails.Height + vbxTrueGrid.Height + 300 Then 'Then
        scrControl.Value = 0
        panDetails.Top = 0
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = panWindow.Width + panWindow.Left
        scrControl.Height = panWindow.Height
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEPERFORM = Nothing  'carmen may 00
    Call NextForm
End Sub

Private Function getJOB()
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset
getJOB = False
On Error GoTo Jobd_Err
If Len(lblJob) > 0 Then
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & lblJob & "'"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsJOB.EOF Then
        clpCode(1).Caption = "Unassigned"
        Exit Function
    End If
    getJOB = True
    clpCode(1).Caption = rsJOB("JB_DESCR")
    rsJOB.Close
End If


Exit Function

Jobd_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "HRJOB", "SELECT")
Call RollBack '28July99 js

End Function


Private Sub medSalary_GotFocus()
 Call SetPanHelp(ActiveControl)
End Sub

Private Sub memComments_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Set_Current_Flag()
Dim SQLQ As String, Msg$
Dim dyn_HRPHHIS As New ADODB.Recordset

On Error GoTo SCFError
If glbMulti Then Exit Sub

'Hemu - 07/07/2003 Begin - Commented out the clone line cause it was giving Error
'                          as 'Row cannot be located for updating'
'Set dyn_HRPHHIS = Data1.Recordset.Clone
dyn_HRPHHIS.Open Data1.RecordSource, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'Hemu- 07/07/2003  End

Screen.MousePointer = HOURGLASS

If dyn_HRPHHIS.RecordCount < 1 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

'Hemu - 07/07/2003 Begin -Check # of records before moving first
If dyn_HRPHHIS.RecordCount > 0 Then dyn_HRPHHIS.MoveFirst
'Hemu - 07/07/2003 End

'EOF?
dyn_HRPHHIS("PH_CURRENT") = True
dyn_HRPHHIS.Update
dyn_HRPHHIS.MoveNext

While Not dyn_HRPHHIS.EOF
    'Hemu - 07/07/2003 Begin - to improve speed, Jaddy suggested
    If dyn_HRPHHIS("PH_CURRENT") <> 0 Then
        dyn_HRPHHIS("PH_CURRENT") = False
        dyn_HRPHHIS.Update
    End If
    'Hemu - 07/07/2003 End
    dyn_HRPHHIS.MoveNext
Wend

dyn_HRPHHIS.Close
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Screen.MousePointer = DEFAULT

Exit Sub

SCFError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PERFORM_HISTORY", "Add")
Call RollBack '28July99 js

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

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'cmdPosition.Enabled = TF
'cmdSalary.Enabled = FT
'vbxTrueGrid.Enabled = FT

chkCurrent.Enabled = TF
'memComments.Enabled = TF
memComments.Locked = FT
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
dlpReviewDate(0).Enabled = TF
dlpReviewDate(1).Enabled = TF
medSalary.Enabled = TF
dlpDate(1).Enabled = TF

'cmdPosCode.Enabled = TF
elpReptAuthShow(0).Enabled = TF
elpReptAuthShow(1).Enabled = TF
elpReptAuthShow(2).Enabled = TF

If glbSetPer Or glbMulti Then
'    cmdPosCode.Visible = False
     clpPosCode.Enabled = TF
Else
     clpPosCode.Enabled = False
End If
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    If glbMulti Then
        clpPosCode.Visible = True
    Else
        clpPosCode.Visible = True 'False 'Hemu - (Linda Approved this) To allow users to enter previous performance for historical positions
    End If
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
Else
    Me.cmdModify_Click
'     clpPosCode.Visible = True
End If

If glbtermopen Then
'    cmdOK.Enabled = False
'    cmdCancel.Enabled = False
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'    cmdPosCode.Visible = False
End If
'If Not gSec_Inq_Salary Then cmdSalary.Enabled = False
'If Not gSec_Inq_Position Then cmdPosition.Enabled = False

'George on Jan 26,2006 #10266
glbDocName = "Performance"
If gsAttachment_DB Then
    'glbJob = "" 'George on Jan 24,2006 #10266
    'glbSDate = "01/01/1900" 'George on Jan 24,2006 #10266
    If Not (rsDATA.BOF And rsDATA.EOF) Then
        'glbJob = rsDATA("PH_JOB")
        'glbSDate = rsDATA("PH_PREVIEW")
        If Not IsNull(rsDATA("PH_DOCKEY")) Then
            glbDocKey = rsDATA("PH_DOCKEY") ' Data1.Recordset("PH_DOCKEY")
        Else
            glbDocKey = 0
        End If
    End If
    Call DispimgIcon(Me, "frmEPERFORM")
    If gSec_Upd_Performance And Not glbtermopen Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If
'George on Jan 26,2006 #10266

End Sub

Private Sub clpPosCode_LostFocus()
Dim X%
On Error GoTo SCError
If orgPosCode = clpPosCode Then Exit Sub
If Not Set_Position(clpPosCode, False) Then Exit Sub

For X% = 0 To 2
    elpReptAuthShow(X%) = ShowEmpnbr(savAuth(X%))
Next
If glbCompSerial <> "S/N - 2368W" Then
    'Ticket #24099 - Show the Current Salary and Effective Date
    'If Set_Salary(clpPosCode, True) Then
    If Set_Salary("", True) Then
        lblCSalary = Round2DEC(fglbCurSalary@)
        lblEDateD = fglbCurSalEdate
    Else
        lblCSalary = ""
        lblEDateD = "Unassigned"
    End If
Else
    lblCSalary = ""
    lblEDateD = ""
End If
Exit Sub
SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "clpPosCode", "HR_JOB_HISTORY", "SELECT")
Call RollBack '

End Sub

Private Sub scrControl_Change()
    panDetails.Top = 0 - scrControl.Value
End Sub

Private Sub txtReptAuthority_Change(Index As Integer)
    elpReptAuthShow(Index).Text = ShowEmpnbr(txtReptAuthority(Index).Text)
End Sub


Private Function updFollow(xType) ' Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsTT As New ADODB.Recordset
Dim Edit1 As Integer

'Ticket #11712, do not create follow up record for TS Tech
If glbCompSerial = "S/N - 2369W" Then
    updFollow = True
    Exit Function
End If

updFollow = False

On Error GoTo CrFollow_Err

If orgNextReviewDate <> "" Then  ' DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'PREV'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(orgNextReviewDate)
   
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If fglbNew And dlpReviewDate(1).Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpReviewDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "PREV"
        
        'Hemu - 04/28/2004 Begin - Jerry said not to pass the comments to followup records
        '                          because of the new privacy of information act.
        '                          Ticket # 6086
        'rsTB("EF_COMMENTS") = memComments
        'Hemu - 04/28/2004 End
        
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='PREV'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "PREV"
            rsTT("TB_DESC") = "Performance Review"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "PREV", "Performance Review")
        
        updFollow = True
        'Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And dlpReviewDate(1).Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpReviewDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "PREV"
        
        'Hemu - 04/28/2004 Begin - Jerry said not to pass the comments to followup records
        '                          because of the new privacy of information act.
        '                          Ticket # 6086
        'rsTB("EF_COMMENTS") = memComments
        'Hemu - 04/28/2004 End
        
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
                
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='PREV'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "PREV"
            rsTT("TB_DESC") = "Performance Review"
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        
        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "PREV", "Performance Review")
        
        updFollow = True
        'Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpReviewDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(dlpReviewDate(1).Text)
            dynHRAT("EF_FREAS") = "PREV"
            
            'Hemu - 04/28/2004 Begin - Jerry said not to pass the comments to followup records
            '                          because of the new privacy of information act.
            '                          Ticket # 6086
            'dynHRAT("EF_COMMENTS") = memComments
            'Hemu - 04/28/2004 End
            
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If orgComments <> memComments Or orgNextReviewDate <> dlpReviewDate(1).Text Then
           ' Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpReviewDate(1).Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
       'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpReviewDate(1).Text = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "Check the Follow up table"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function
'Private Sub txtReviewDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, X As Integer
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value
Call SET_UP_MODE
If Not Data1.Recordset.EOF Then
    

    'Ticket #24099 - Moved the line below down after getting the current position. If for some reason the employee
    'do not have that position anymore then the Current Salary is not getting retrieved.
    'If Not Set_Position(clpPosCode.Text, False) Then Exit Sub
    If glbCompSerial <> "S/N - 2368W" Then
        'Ticket #24099 - Show the Current Salary and Effective Date
        'If Set_Salary(clpPosCode.Text, True) Then
        If Set_Salary("", True) Then
            lblCSalary = Round2DEC(fglbCurSalary@)
            lblEDateD = fglbCurSalEdate
        Else
            lblCSalary = ""
            lblEDateD = ""
        End If
    End If
    If Not Set_Position(clpPosCode.Text, False) Then Exit Sub
Else
    'lblJobDesc.Caption = "Unassigned"
    lblCSalary = ""
    lblEDateD = ""
End If

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_PERFORM_HISTORY", "Add")
Resume Next

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

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo JobHis_Err

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

Private Sub CR_SalHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo SalHis_Err

Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_SDATE DESC"

    If dynaSalHIS.State <> 0 Then dynaSalHIS.Close
    dynaSalHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE SH_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_SDATE DESC"
    
    If dynaSalHIS.State <> 0 Then dynaSalHIS.Close
    dynaSalHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If
Screen.MousePointer = DEFAULT

Exit Sub

SalHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_SALARY_HISTORY", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub


Private Function Set_Position(nJob As String, nCurrent As Boolean)
Dim SQLQ As String, Msg$

Set_Position = False
On Error GoTo SCError
Screen.MousePointer = HOURGLASS
dynaJobHIS.Requery

SQLQ = ""
If nCurrent Then SQLQ = SQLQ & " JH_CURRENT<>0 "
If nJob <> "" Then SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_JOB='" & nJob & "' "
dynaJobHIS.Filter = SQLQ

If dynaJobHIS.BOF And dynaJobHIS.EOF Then
    glbStopPerform% = nCurrent
    Screen.MousePointer = DEFAULT
    dynaJobHIS.Filter = ""
    Exit Function
Else
    glbStopPerform% = False
End If

If Not IsNull(dynaJobHIS("JH_JOB")) Then lblJob.Caption = dynaJobHIS("JH_JOB") Else lblJob.Caption = ""    ' record
If IsNull(dynaJobHIS("JH_ID")) Then fglbJobID& = 0 Else fglbJobID& = dynaJobHIS("JH_ID")
If IsNull(dynaJobHIS("JH_REPTAU")) Then savAuth(0) = "" Else savAuth(0) = dynaJobHIS("JH_REPTAU")
If IsNull(dynaJobHIS("JH_REPTAU2")) Then savAuth(1) = "" Else savAuth(1) = dynaJobHIS("JH_REPTAU2")
If IsNull(dynaJobHIS("JH_REPTAU3")) Then savAuth(2) = "" Else savAuth(2) = dynaJobHIS("JH_REPTAU3")
dynaJobHIS.Filter = ""
Set_Position = True
Screen.MousePointer = DEFAULT
Exit Function

SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_HISTORY", "SELECT")
Call RollBack '28July99 js

End Function


Private Function Set_Salary(nJob As String, nCurrent As Boolean)
Dim SQLQ As String, Msg$

Set_Salary = False
On Error GoTo SCError
Screen.MousePointer = HOURGLASS
dynaSalHIS.Requery
SQLQ = ""
If nCurrent Then SQLQ = SQLQ & " SH_CURRENT<>0 "
If nJob <> "" Then SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " SH_JOB='" & nJob & "' "
dynaSalHIS.Filter = SQLQ

If dynaSalHIS.BOF And dynaSalHIS.EOF Then
    glbStopPerform% = nCurrent
    Screen.MousePointer = DEFAULT
    dynaSalHIS.Filter = ""
    Exit Function
Else
    glbStopPerform% = False
End If

fglbCurSalary@ = dynaSalHIS("SH_SALARY")
fglbCurSalEdate = dynaSalHIS("SH_EDATE")

dynaSalHIS.Filter = ""
Set_Salary = True
Screen.MousePointer = DEFAULT
Exit Function

SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_Salary_HISTORY", "SELECT")
Call RollBack '28July99 js

End Function

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
strNUM = "0." & String(glbCompDecHR, "0")
Round2DEC = Format(Round(tmpNUM, glbCompDecHR), strNUM)

If glbWFC And locCountry = "AUSTRALIA" Then
    Round2DEC = Round(tmpNUM, 4)
End If
End Function

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    Dim X, xFld
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
                rsDATA.CursorLocation = adUseServer
            End If
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If gsAttachment_DB Then
            If rsDATA.EOF Or rsDATA.BOF Then
                imgSec.Visible = False
                imgNoSec.Visible = True
            End If
        End If
    Else
        If glbtermopen Then
            SQLQ = "SELECT Term_PERFORM_HISTORY.*,"
        Else
            SQLQ = "SELECT HR_PERFORM_HISTORY.*,"
        End If
        For X = 0 To 2
            xFld = "REPTAU" & IIf(X = 0, "", X + 1)
            If glbLinamar Then
                SQLQ = SQLQ & " CASE WHEN PH_" & xFld & " IS NOT NULL AND LEN(PH_" & xFld & ")>2 "
                SQLQ = SQLQ & " THEN RIGHT(PH_" & xFld & ",3)+'-'+"
                SQLQ = SQLQ & " LEFT(PH_" & xFld & ",LEN(PH_" & xFld & ")-3) "
                SQLQ = SQLQ & " ELSE STR(PH_" & xFld & ") END "
                SQLQ = SQLQ & " AS " & xFld & IIf(X = 2, "", ",")
            Else
                If glbOracle Then
                    SQLQ = SQLQ & "PH_" & xFld & " AS " & xFld & IIf(X = 2, "", ",")
                Else
                    SQLQ = SQLQ & "STR(PH_" & xFld & ") AS " & xFld & IIf(X = 2, "", ",")
                End If
            End If
        Next
        If glbtermopen Then
            SQLQ = SQLQ & " FROM Term_PERFORM_HISTORY "
            SQLQ = SQLQ & " WHERE PH_ID=" & Data1.Recordset!PH_ID
            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = SQLQ & " FROM  HR_PERFORM_HISTORY"
            SQLQ = SQLQ & " WHERE PH_ID = " & Data1.Recordset!PH_ID
            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
                rsDATA.CursorLocation = adUseServer
            End If
            rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If gsAttachment_DB Then
            If rsDATA.EOF Or rsDATA.BOF Then
                imgSec.Visible = False
                imgNoSec.Visible = True
            End If
        End If

        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
        Call Set_Control("R", Me, rsDATA)
        If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
            chkCurrent = Data1.Recordset("PH_CURRENT")
        End If
    End If
    
'Comment by Frank on 05/16/07, it will do orgNextReviewDate = dlpReviewDate(1).Text on cmdModify_Click
'before updFollow("U")
'Call SET_UP_MODE
'    Me.cmdModify_Click
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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Performance
End Property

Public Property Get Addable() As Boolean
Addable = Not glbtermopen
End Property
Public Property Get Updateble() As Boolean

Updateble = Not glbtermopen
End Property
Public Property Get Deleteble() As Boolean

Deleteble = Not glbtermopen
End Property
Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
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
If Not UpdateRight Then TF = False
If Not Updateble Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEPERFORM.Caption = lStr("Performance") & " - " & Left$(glbLEE_SName, 5)
    frmEPERFORM.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
 If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub


Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEPERFORM")
    Call FillMemoFile(SQLQ, "Performance")
End Sub

Private Sub cmdImport_Click()
    'Ticket #29703
    If fglbNew Then
        If Len(Trim(clpPosCode.Text)) = 0 Then
            MsgBox "Position is mandatory to attach a document", vbInformation, "Invalid Position"
            Exit Sub
        End If
        glbJob = clpPosCode.Text
    Else
        glbJob = rsDATA("PH_JOB")
    End If
    
    glbDocNewRecord = fglbNew
    glbDocName = "Performance"
    If fglbNew Then
        glbDocKey = 0
    Else
        'Ticket #23969 - for some records the PH_ID is not same as the _DOCKEY so it's failing to see that doc is attached
        'glbDocKey = rsDATA("PH_ID")
        If IsNull(rsDATA("PH_DOCKEY")) Then
            glbDocKey = rsDATA("PH_ID")
        Else
            glbDocKey = rsDATA("PH_DOCKEY")
        End If
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEPERFORM")
End Sub


Private Sub Updstats_Change(Index As Integer)
    If Index = 0 And Not glbWFC Then
        'dlpDate(2).Text = Updstats(0)
    End If
    If Index = 2 Then
        lblUserDesc = GetUserDesc(Updstats(2))
    End If
End Sub

Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
    'If gsEMAIL_ONPERFORMANCE Then
    'Ticket #22409 Frank 08/08/2012 - add "not glbWFC" since they use this email for Smoker Status Change
    If gsEMAIL_ONPERFORMANCE And Not glbWFC Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
        
        'Ticket #20317 - Send email to More Emails list as well.
        xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
        End If
            
        'If Len(xEmail) > 0 Then    'Hemu - (Ticket #11562) - Jerry asked to remove the check for email address presence.
            frmSendEmail.txtTo.Text = xToEmail
            If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352, do not cc it to employee
            Else
                frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            End If
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR " & lStr("Performance") & " Change Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            'End If
        '    MsgBox "There is no email address for the 'Email Notification on " & lstr("Performance") & " ' on Company Preference screen. "
        'End If


    End If
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

Private Function isReviewDatesChanged() 'Ticket #21601 Franks 02/24/2012
Dim retval As Boolean
    retval = False
    If chkCurrent.Value Then 'current only
        'If IsDate(orgReviewDate) Or IsDate(orgNextReviewDate) Then 'not the first record
        If Len(orgReviewDate) = 0 And Len(orgNextReviewDate) = 0 Then 'first record
            retval = True
        Else
            'Review Date
            If Len(orgReviewDate) = 0 Then
                If IsDate(dlpReviewDate(0).Text) Then
                    retval = True
                End If
            Else
                If IsDate(orgReviewDate) And IsDate(dlpReviewDate(0).Text) Then
                    If Not CVDate(orgReviewDate) = CVDate(dlpReviewDate(0).Text) Then
                        retval = True
                    End If
                End If
            End If
            'Next Review Date
            If Len(orgNextReviewDate) = 0 Then
                If IsDate(dlpReviewDate(1).Text) Then
                    retval = True
                End If
            Else
                If IsDate(orgNextReviewDate) And IsDate(dlpReviewDate(1).Text) Then
                    If Not CVDate(orgNextReviewDate) = CVDate(dlpReviewDate(1).Text) Then
                        retval = True
                    End If
                End If
            End If
        End If
        'End If
    End If
    isReviewDatesChanged = retval
End Function

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    'If gsEMAIL_ONPOSITION Then
        If Not UserEmailExist Then
            Exit Sub
        End If

        xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
        End If
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
            If Len(xBranch) > 0 Then
                xBranch = GetTABLDesc("EDSE", xBranch)
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR Performance Change Notice - " & xBranch & lblEEName.Caption
            frmSendEmail.txtSubject.Text = xEmailSubject
        
            frmSendEmail.txtBody.Text = MailBody
            'frmSendEmail.Show 1
            MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(0).Caption = "Sending email..."
            frmSendEmail.Tag = ""
            frmSendEmail.cmdSend_Click
            Do
                DoEvents
            Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
            ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
            If frmSendEmail.Tag = "DONE" Then
                Unload frmSendEmail
            Else
                Unload frmSendEmail
            End If
            MDIMain.panHelp(0).Caption = ""
            MDIMain.panHelp(0).FloodType = 1
        End If

    'End If
    Exit Sub

Email_Err:
    'If Err.Number = 364 Then
    '    Exit Sub
    'End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "EmailSendingForSamuel")
    'Resume Next
    Exit Sub

End Sub

Private Function CheckDuplCurrent(xEmpNo, xJobCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retval As Boolean
    retval = False
    SQLQ = "SELECT * FROM HR_PERFORM_HISTORY "
    SQLQ = SQLQ & " WHERE PH_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & " AND PH_JOB = '" & xJobCode & "' "
    SQLQ = SQLQ & " AND PH_CURRENT <>0 " 'current
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = True
    End If
    rsTemp.Close
    CheckDuplCurrent = retval
End Function

'Private Function Older_FollowUp_Records_Found() As Boolean
'    Dim rsFollowUp As New ADODB.Recordset
'    Dim SQLQ As String
'
'    SQLQ = "SELECT * FROM HR_FOLLOW_UP "
'    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND EF_FREAS = 'PREV'"   'SREV, PREV, EDUC
'    SQLQ = SQLQ & " AND EF_COMPLETED <> 1"  'Not completed
'    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsFollowUp.EOF Then
'        Older_FollowUp_Records_Found = True
'    Else
'        Older_FollowUp_Records_Found = False
'    End If
'End Function

Private Function NextReviewDate_Macaulay()
    Dim rsJOB As New ADODB.Recordset
    Dim SQLQ As String
    Dim xHoursWeek
    Dim xNoOfWeeks
    Dim xNoOfWorkDays
    Dim xPosStartDate
    Dim xRevDate
    Dim xStartDate
    
    'Check if the Performance Review code is correct, e.g. 999x
    If Len(clpCode(1).Text) > 0 Then
        If IsNumeric(Left(clpCode(1).Text, Len(clpCode(1).Text) - 1)) And Not IsNumeric(Right(clpCode(1).Text, 1)) Then
            xStartDate = ""
            SQLQ = ""
            xHoursWeek = ""
            
            'Day or Hours then only proceed
            If Right(clpCode(1).Text, 1) = "D" Then
                'Compute Next Review Date by Days
                'Retrieve employee's Position Start Date based on the Job ID associated with this
                'Performance record's Position if this the first Performance Review record else go by Current Review Date
                If Data1.Recordset.EOF Then
                    'This is first Performance Review record use Position Start Date as Start Date
                    SQLQ = "SELECT JH_SDATE, JH_EMPNBR, JH_WHRS FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & glbLEE_ID
                    SQLQ = SQLQ & " AND JH_JOB = '" & clpPosCode.Text & "'"
                    'If IsNumeric(lblJobID) Then
                    '    SQLQ = SQLQ & " AND JH_ID = " & lblJobID
                    'End If
                    SQLQ = SQLQ & " ORDER BY JH_CURRENT DESC, JH_SDATE DESC"
                    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsJOB.EOF Then
                        rsJOB.MoveFirst
                        xStartDate = rsJOB("JH_SDATE")
                    End If
                    rsJOB.Close
                    Set rsJOB = Nothing
                Else
                    'This is not the first Performance Review record use Current Review Date as Start Date.
                    If IsDate(dlpReviewDate(0).Text) Then
                        xStartDate = dlpReviewDate(0).Text
                    Else
                        xStartDate = ""
                    End If
                End If
            ElseIf Right(clpCode(1).Text, 1) = "H" Then
                'Compute Next Review Date by Hours
                'Retrive employee's Position Hours/Week based on the Job ID associated with this
                'Performance record's Position if this the first Performance Review record else go by Current Review Date
                
                'Retrieve Job record to get Hours per Weeks anyways independent of if this is first Perf. Review
                'record or not
                SQLQ = "SELECT JH_SDATE, JH_EMPNBR, JH_WHRS FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & glbLEE_ID
                SQLQ = SQLQ & " AND JH_JOB = '" & clpPosCode.Text & "'"
                'If IsNumeric(lblJobID) Then
                '    SQLQ = SQLQ & " AND JH_ID = " & lblJobID
                'End If
                SQLQ = SQLQ & " ORDER BY JH_CURRENT DESC, JH_SDATE DESC"
                rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsJOB.EOF Then
                    rsJOB.MoveFirst
                    xStartDate = rsJOB("JH_SDATE")
                    
                    'if valid Hours/Week
                    If IsNumeric(rsJOB("JH_WHRS")) And rsJOB("JH_WHRS") <> 0 Then
                        xHoursWeek = rsJOB("JH_WHRS")
                    Else
                        xHoursWeek = ""
                    End If
                End If
                rsJOB.Close
                Set rsJOB = Nothing
                                
                'If this is not the first Performance Review record then use Current Review Date as Start Date.
                If Not Data1.Recordset.EOF Then
                    If IsDate(dlpReviewDate(0).Text) Then
                        xStartDate = dlpReviewDate(0).Text
                    Else
                        xStartDate = ""
                    End If
                End If
            End If
                    
            If Right(clpCode(1).Text, 1) = "D" Then
                'If valid Start Date
                If IsDate(xStartDate) Then
                    'Add # Days to the Start Date excluding Weekends and Statutory Holidays
                    xRevDate = AddWorkingDays(xStartDate, Left(clpCode(1).Text, Len(clpCode(1).Text) - 1), True)
                    
                    'Return the Date
                    dlpReviewDate(1).Text = xRevDate
                End If
            ElseIf Right(clpCode(1).Text, 1) = "H" Then
                '# of Weeks and # of Days
                xNoOfWeeks = 0
                xNoOfWorkDays = 0
                If IsNumeric(xHoursWeek) And xHoursWeek <> 0 Then
                    xNoOfWeeks = Left(clpCode(1).Text, Len(clpCode(1).Text) - 1) / xHoursWeek
                    
                    'Convert # of Weeks to # of Days
                    xNoOfWorkDays = xNoOfWeeks * 5
                    
                    'Add # of Hours (in Days now) to the Start Date excluding Weekends and Statutory Holidays
                    xRevDate = AddWorkingDays(xStartDate, xNoOfWorkDays, True)
                    
                    'Return the Date
                    dlpReviewDate(1).Text = xRevDate
                End If
            End If
        End If
    End If
End Function
