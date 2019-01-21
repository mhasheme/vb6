VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmESALARY 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Salary History"
   ClientHeight    =   10950
   ClientLeft      =   195
   ClientTop       =   1200
   ClientWidth     =   12840
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
   ScaleWidth      =   12840
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   7215
      LargeChange     =   300
      Left            =   11760
      Max             =   4000
      SmallChange     =   300
      TabIndex        =   105
      Top             =   2280
      Width           =   350
   End
   Begin Threed.SSPanel panWindow 
      Height          =   7335
      Left            =   0
      TabIndex        =   39
      Top             =   2160
      Width           =   11655
      _Version        =   65536
      _ExtentX        =   20558
      _ExtentY        =   12938
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
      Begin VB.TextBox txtVadPayRate 
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
         Left            =   5280
         MaxLength       =   25
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   6960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtVadSalRate 
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
         Left            =   5880
         MaxLength       =   25
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   6960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtVadAddModDel 
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
         Left            =   6480
         MaxLength       =   25
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   6960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox panDetails 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7095
         Left            =   0
         ScaleHeight     =   7095
         ScaleWidth      =   11655
         TabIndex        =   40
         Top             =   0
         Width           =   11655
         Begin VB.TextBox txtPosGroup 
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
            Left            =   2060
            TabIndex        =   113
            Tag             =   "Position Group"
            Top             =   660
            Visible         =   0   'False
            Width           =   3360
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataSource      =   " "
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   89
            Tag             =   "00-Enter Union Code"
            Top             =   7200
            Visible         =   0   'False
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin VB.TextBox txtWHRS 
            Appearance      =   0  'Flat
            DataField       =   "SH_WHRS"
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
            Left            =   4920
            MaxLength       =   25
            TabIndex        =   76
            TabStop         =   0   'False
            Tag             =   "10-Hours per Week"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Updstats 
            Appearance      =   0  'Flat
            DataField       =   "SH_LDATE"
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   1620
            MaxLength       =   25
            TabIndex        =   75
            TabStop         =   0   'False
            Text            =   "Ldate"
            Top             =   7050
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Updstats 
            Appearance      =   0  'Flat
            DataField       =   "SH_LTIME"
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   1
            Left            =   270
            MaxLength       =   25
            TabIndex        =   74
            TabStop         =   0   'False
            Text            =   "LTime"
            Top             =   7050
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Updstats 
            Appearance      =   0  'Flat
            DataField       =   "SH_LUSER"
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
            Left            =   2460
            MaxLength       =   25
            TabIndex        =   73
            TabStop         =   0   'False
            Text            =   "LUser"
            Top             =   7020
            Visible         =   0   'False
            Width           =   975
         End
         Begin Threed.SSFrame fraSalary 
            Height          =   1515
            Left            =   60
            TabIndex        =   77
            Top             =   2250
            Width           =   9525
            _Version        =   65536
            _ExtentX        =   16801
            _ExtentY        =   2672
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   1
            Begin VB.TextBox txtVStep 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   7320
               MaxLength       =   10
               TabIndex        =   79
               TabStop         =   0   'False
               Tag             =   "01-Country"
               Top             =   1200
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.TextBox txtVGroup 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
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
               Left            =   7320
               MaxLength       =   10
               TabIndex        =   78
               TabStop         =   0   'False
               Top             =   720
               Visible         =   0   'False
               Width           =   1155
            End
            Begin VB.ComboBox cboVStep 
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
               Left            =   5280
               TabIndex        =   24
               Top             =   1170
               Width           =   2055
            End
            Begin VB.ComboBox cboVGRoup 
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
               Left            =   5280
               TabIndex        =   23
               Top             =   735
               Width           =   2055
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
               Left            =   5400
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Tag             =   "01-Choose annum or hour"
               Top             =   180
               Width           =   1215
            End
            Begin VB.ComboBox comSalScale 
               Height          =   315
               Left            =   7620
               TabIndex        =   19
               Tag             =   "00-Position Grid Steps"
               Top             =   180
               Width           =   675
            End
            Begin MSMask.MaskEdBox medsalary 
               DataField       =   "SH_SALARY"
               Height          =   285
               Left            =   1665
               TabIndex        =   16
               Tag             =   "21-Enter salary"
               Top             =   195
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
            Begin MSMask.MaskEdBox medPremium 
               Height          =   285
               Left            =   1665
               TabIndex        =   21
               Tag             =   "21-Enter salary"
               Top             =   750
               Width           =   1290
               _ExtentX        =   2275
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
            Begin MSMask.MaskEdBox medTotal 
               Height          =   285
               Left            =   1665
               TabIndex        =   22
               Tag             =   "21-Enter salary"
               Top             =   1185
               Width           =   1290
               _ExtentX        =   2275
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
               Height          =   285
               Index           =   6
               Left            =   3720
               TabIndex        =   17
               Tag             =   "00-Currency Indicator - Code"
               Top             =   195
               Visible         =   0   'False
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               ShowUnassigned  =   1
               TABLName        =   "WFCI"
               Enabled         =   0   'False
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Per"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   6
               Left            =   5040
               TabIndex        =   83
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lblCurrencyIndicator 
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
               Left            =   3000
               TabIndex        =   118
               Top             =   240
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.Label lblPosGrp 
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
               Left            =   7080
               TabIndex        =   115
               Top             =   240
               Visible         =   0   'False
               Width           =   45
            End
            Begin VB.Label lblTitle 
               Caption         =   "Vailtech Step"
               Height          =   255
               Index           =   20
               Left            =   3480
               TabIndex        =   88
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label lblTitle 
               Caption         =   "Vailtech Group"
               Height          =   255
               Index           =   19
               Left            =   3480
               TabIndex        =   87
               Top             =   765
               Width           =   1455
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Step"
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
               Left            =   6960
               TabIndex        =   86
               Top             =   240
               Width           =   600
            End
            Begin VB.Label lblSalaryGrade 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "SH_GRADE"
               DataField       =   "SH_GRADE"
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
               Left            =   8400
               TabIndex        =   85
               Top             =   240
               Visible         =   0   'False
               Width           =   885
            End
            Begin VB.Label lblSalCode 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "SalCode"
               DataField       =   "SH_SALCD"
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
               Left            =   6240
               TabIndex        =   84
               Top             =   480
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblTitle 
               Caption         =   "Total"
               Height          =   255
               Index           =   18
               Left            =   120
               TabIndex        =   82
               Top             =   1200
               Width           =   1455
            End
            Begin VB.Label lblTitle 
               Caption         =   "Premium"
               Height          =   255
               Index           =   16
               Left            =   120
               TabIndex        =   81
               Top             =   765
               Width           =   1455
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Salary"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   5
               Left            =   150
               TabIndex        =   80
               Top             =   240
               Width           =   1380
            End
         End
         Begin INFOHR_Controls.DateLookup dlpPosStDate 
            DataField       =   "SH_SDATE"
            Height          =   285
            Left            =   1740
            TabIndex        =   2
            Tag             =   "41-Enter Position Start Date"
            Top             =   330
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            ShowDescription =   0   'False
            TextBoxWidth    =   1180
         End
         Begin INFOHR_Controls.CodeLookup clpPostCode 
            DataField       =   "SH_JOB"
            Height          =   285
            Left            =   1740
            TabIndex        =   1
            Tag             =   "01-Position code"
            Top             =   0
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   5
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "SH_SREAS1"
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Tag             =   "01-Reason code "
            Top             =   1290
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDRC"
         End
         Begin Threed.SSCheck chkCurrent 
            DataField       =   "SH_CURRENT"
            Height          =   255
            Left            =   6900
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   60
            Width           =   1890
            _Version        =   65536
            _ExtentX        =   3334
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Current Salary Record"
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
         Begin MSMask.MaskEdBox medPercentChng 
            DataField       =   "SH_SALPC1"
            Height          =   285
            Index           =   1
            Left            =   5310
            TabIndex        =   8
            Tag             =   "10-Percentage change from previous salary"
            Top             =   1290
            Width           =   870
            _ExtentX        =   1535
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
            Format          =   "0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPercentChng 
            DataField       =   "SH_SALPC2"
            Height          =   285
            Index           =   2
            Left            =   5310
            TabIndex        =   11
            Tag             =   "10-Percentage change from previous salary"
            Top             =   1605
            Width           =   870
            _ExtentX        =   1535
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
            Format          =   "0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medPercentChng 
            DataField       =   "SH_SALPC3"
            Height          =   285
            Index           =   3
            Left            =   5310
            TabIndex        =   14
            Tag             =   "10-Percentage change from previous salary"
            Top             =   1920
            Width           =   870
            _ExtentX        =   1535
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
            Format          =   "0.00%"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medAmtChng 
            DataField       =   "SH_SALCHG1"
            Height          =   285
            Index           =   1
            Left            =   7230
            TabIndex        =   9
            Tag             =   "20-$ change from previous salary"
            Top             =   1290
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medAmtChng 
            DataField       =   "SH_SALCHG2"
            Height          =   285
            Index           =   2
            Left            =   7230
            TabIndex        =   12
            Tag             =   "20-$ change from previous salary"
            Top             =   1605
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox medAmtChng 
            DataField       =   "SH_SALCHG3"
            Height          =   285
            Index           =   3
            Left            =   7230
            TabIndex        =   15
            Tag             =   "20-$ change from previous salary"
            Top             =   1920
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   "$#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "SH_SREAS2"
            Height          =   285
            Index           =   2
            Left            =   240
            TabIndex        =   10
            Tag             =   "01-Reason code "
            Top             =   1620
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDRC"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "SH_SREAS3"
            Height          =   285
            Index           =   3
            Left            =   240
            TabIndex        =   13
            Tag             =   "01-Reason code "
            Top             =   1950
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "SDRC"
         End
         Begin MSMask.MaskEdBox medNFacSalary 
            DataField       =   "SH_NFAC_SALARY"
            Height          =   285
            Left            =   1720
            TabIndex        =   20
            Tag             =   "21-Enter non-factored salary"
            Top             =   2880
            Visible         =   0   'False
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
         Begin INFOHR_Controls.CodeLookup clpGrid 
            DataField       =   "SH_GRID"
            Height          =   285
            Left            =   1740
            TabIndex        =   3
            Top             =   660
            Visible         =   0   'False
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "JBGD"
            TABLTitle       =   "Grid Category"
            MaxLength       =   10
         End
         Begin Threed.SSCheck chkRedCircled 
            Height          =   255
            Left            =   9240
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   1290
            _Version        =   65536
            _ExtentX        =   2275
            _ExtentY        =   450
            _StockProps     =   78
            Caption         =   "Red-Circled"
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
         Begin VB.Frame fraSalary2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2895
            Left            =   60
            TabIndex        =   41
            Top             =   3750
            Width           =   11415
            Begin VB.TextBox txtFiscalYear 
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
               Left            =   9660
               MaxLength       =   4
               TabIndex        =   48
               Tag             =   "00-Fiscal Year"
               Top             =   435
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtPayrollID 
               Appearance      =   0  'Flat
               DataField       =   "SH_PAYROLL_ID"
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
               Left            =   6300
               MaxLength       =   25
               TabIndex        =   47
               Tag             =   "00-Payroll ID"
               Top             =   420
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.OptionButton optUserSys 
               Caption         =   "User"
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
               Left            =   3240
               TabIndex        =   46
               Top             =   1740
               Width           =   1095
            End
            Begin VB.OptionButton optUserSys 
               Caption         =   "System"
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
               Left            =   2280
               TabIndex        =   45
               Top             =   1740
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.TextBox txtMarketLine 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               DataField       =   "SH_MarketLine"
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
               Height          =   255
               Left            =   6330
               TabIndex        =   44
               Top             =   810
               Visible         =   0   'False
               Width           =   850
            End
            Begin VB.ComboBox cmbMarketLine 
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
               Left            =   6300
               Style           =   2  'Dropdown List
               TabIndex        =   43
               Tag             =   "00-Market Line"
               Top             =   780
               Width           =   1155
            End
            Begin VB.TextBox txtUserSys 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               DataField       =   "SH_COMPA_USER"
               Height          =   285
               Left            =   3570
               TabIndex        =   42
               Top             =   1770
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtComment 
               Appearance      =   0  'Flat
               DataField       =   "SH_COMMENT"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   1800
               MaxLength       =   100
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   26
               Tag             =   "00-Position Comments"
               Top             =   885
               Width           =   3285
            End
            Begin INFOHR_Controls.DateLookup dlpDate 
               DataField       =   "SH_EDATE"
               Height          =   285
               Index           =   0
               Left            =   1485
               TabIndex        =   25
               Tag             =   "41-Effective date of salary change"
               Top             =   420
               Width           =   2580
               _ExtentX        =   4551
               _ExtentY        =   503
               TextBoxWidth    =   1215
            End
            Begin INFOHR_Controls.DateLookup dlpDate 
               DataField       =   "SH_NEXTDAT"
               Height          =   285
               Index           =   1
               Left            =   6300
               TabIndex        =   28
               Tag             =   "40-Next Date to Review Salary"
               Top             =   2130
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   503
               TextBoxWidth    =   1215
            End
            Begin INFOHR_Controls.CodeLookup clpCode 
               DataField       =   "SH_PAYP"
               Height          =   285
               Index           =   4
               Left            =   1440
               TabIndex        =   27
               Tag             =   "00-Enter pay period code"
               Top             =   2100
               Width           =   3405
               _ExtentX        =   6006
               _ExtentY        =   503
               ShowUnassigned  =   1
               TABLName        =   "SDPP"
            End
            Begin MSMask.MaskEdBox mskCampa 
               Height          =   285
               Left            =   4080
               TabIndex        =   49
               Top             =   1740
               Width           =   1095
               _ExtentX        =   1931
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
            Begin INFOHR_Controls.CodeLookup clpCode 
               Height          =   285
               Index           =   0
               Left            =   8325
               TabIndex        =   50
               Tag             =   "00-Section - Code"
               Top             =   120
               Visible         =   0   'False
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   503
               ShowUnassigned  =   1
               TABLName        =   "EDSE"
            End
            Begin INFOHR_Controls.DateLookup dlpDate 
               DataField       =   "SH_TRANSDATE"
               Height          =   285
               Index           =   2
               Left            =   6300
               TabIndex        =   29
               Tag             =   "40-Transaction Date"
               Top             =   2490
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   503
               TextBoxWidth    =   1215
               Enabled         =   0   'False
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  'Right Justify
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
               Left            =   4755
               TabIndex        =   72
               Top             =   2535
               Width           =   1455
            End
            Begin VB.Label txtLambtonJob 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
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
               Left            =   6330
               TabIndex        =   71
               Top             =   840
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Label lblLambtonJob 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Occupation"
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
               TabIndex        =   70
               Top             =   870
               Visible         =   0   'False
               Width           =   825
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
               Left            =   1440
               TabIndex        =   69
               Top             =   2505
               Width           =   2775
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
               Left            =   120
               TabIndex        =   68
               Top             =   2505
               Width           =   1095
            End
            Begin VB.Label lblPlant 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Plant "
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   7440
               TabIndex        =   67
               Top             =   120
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblFiscalYear 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Fiscal Year"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   8640
               TabIndex        =   66
               Top             =   450
               Visible         =   0   'False
               Width           =   960
            End
            Begin VB.Label lblPayID 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
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
               Left            =   5280
               TabIndex        =   65
               Top             =   420
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Salary Scale"
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
               Index           =   13
               Left            =   5280
               TabIndex        =   64
               Top             =   1170
               Width           =   960
            End
            Begin VB.Label lblMLine 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BackStyle       =   0  'Transparent
               Caption         =   "Market Line"
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
               Left            =   7560
               TabIndex        =   63
               Top             =   840
               Width           =   840
            End
            Begin VB.Label lblMarketLine 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Market Line"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   5280
               TabIndex        =   62
               Top             =   810
               Width           =   1020
            End
            Begin VB.Label lblsalstate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   2
               Left            =   8220
               TabIndex        =   61
               Top             =   1170
               Width           =   885
            End
            Begin VB.Label lblsalstate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   1
               Left            =   7260
               TabIndex        =   60
               Top             =   1170
               Width           =   885
            End
            Begin VB.Label lblsalstate 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
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
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   6300
               TabIndex        =   59
               Top             =   1170
               Width           =   885
            End
            Begin VB.Label lblComment 
               AutoSize        =   -1  'True
               Caption         =   "SComments"
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
               TabIndex        =   58
               Top             =   900
               Width           =   840
            End
            Begin VB.Label lblWhrs 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               DataField       =   "SH_WHRS"
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
               Left            =   6690
               TabIndex        =   57
               Top             =   150
               Visible         =   0   'False
               Width           =   90
            End
            Begin VB.Label lblTitle 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Next Review"
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
               Left            =   5295
               TabIndex        =   56
               Top             =   2160
               Width           =   915
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Hours per Week"
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
               Index           =   7
               Left            =   5280
               TabIndex        =   55
               Top             =   150
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Pay Period"
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
               TabIndex        =   54
               Top             =   2130
               Width           =   1365
            End
            Begin VB.Label lblCompaNum 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               DataField       =   "SH_COMPA"
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
               Left            =   1890
               TabIndex        =   53
               Top             =   1770
               Width           =   90
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Compa-Ratio"
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
               Left            =   120
               TabIndex        =   52
               Top             =   1740
               Width           =   1095
            End
            Begin VB.Label lblTitle 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Effective Date"
               ForeColor       =   &H80000008&
               Height          =   195
               Index           =   8
               Left            =   120
               TabIndex        =   51
               Top             =   420
               Width           =   1245
            End
         End
         Begin INFOHR_Controls.CodeLookup clpDiv 
            Height          =   285
            Left            =   1740
            TabIndex        =   4
            Tag             =   "00-Specific Division Desired"
            Top             =   660
            Visible         =   0   'False
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   1
         End
         Begin Threed.SSCheck chkPrimary 
            DataField       =   "SH_PRIMARY"
            Height          =   285
            Left            =   9240
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   78
            Caption         =   "Primary Position"
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
            Left            =   0
            TabIndex        =   116
            Top             =   720
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblPosGroup 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Group"
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
            Left            =   0
            TabIndex        =   114
            Top             =   690
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblSalLevel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Level"
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
            TabIndex        =   112
            Top             =   6720
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.Label lblNFacSal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Non-Factored Salary"
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   180
            TabIndex        =   106
            Top             =   2925
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblHoursPay 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8460
            TabIndex        =   104
            Top             =   705
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hourly Rate :"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   21
            Left            =   6900
            TabIndex        =   103
            Top             =   705
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblBANDCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disp"
            DataField       =   "SH_Band"
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
            Left            =   10380
            TabIndex        =   102
            Top             =   2520
            Width           =   315
         End
         Begin VB.Label lblBand 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Band"
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
            Left            =   9780
            TabIndex        =   101
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblPayPeriodSalary 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   8460
            TabIndex        =   99
            Top             =   405
            Width           =   75
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Per Pay :"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   6900
            TabIndex        =   98
            Top             =   405
            Width           =   1455
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours per Week:"
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
            Left            =   3540
            TabIndex        =   97
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Amount Change"
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
            Left            =   6990
            TabIndex        =   96
            Top             =   1050
            Width           =   1350
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Percentage Change"
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
            Left            =   5070
            TabIndex        =   95
            Top             =   1050
            Width           =   1695
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason For Salary Change"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   94
            Top             =   1050
            Width           =   2280
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Start Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   93
            Top             =   390
            Width           =   1620
         End
         Begin VB.Label lblEEID 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "EEId"
            DataField       =   "SH_EMPNBR"
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
            Left            =   1140
            TabIndex        =   92
            Top             =   7170
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lblCNum 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comp"
            DataField       =   "SH_COMPNO"
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
            Left            =   1260
            TabIndex        =   91
            Top             =   6840
            Visible         =   0   'False
            Width           =   405
         End
         Begin VB.Label LabelPos 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   90
            Top             =   60
            Width           =   765
         End
         Begin VB.Label lblGrid 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Grid Category"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   100
            Top             =   690
            Visible         =   0   'False
            Width           =   1170
         End
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   34
      Top             =   10530
      Width           =   12840
      _Version        =   65536
      _ExtentX        =   22648
      _ExtentY        =   741
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
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdTranDate 
         Appearance      =   0  'Flat
         Caption         =   "Edit Transaction Date"
         Height          =   375
         Left            =   4680
         TabIndex        =   107
         Tag             =   "Edit Transaction Date"
         Top             =   0
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmdChPos 
         Appearance      =   0  'Flat
         Caption         =   "Edit Position/&Date"
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Tag             =   "Edit Position Code and Start Date"
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmdRecal 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate"
         Height          =   375
         Left            =   2520
         TabIndex        =   36
         Tag             =   "Recalculate Percentage Change"
         Top             =   0
         Width           =   1815
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   8250
         Top             =   30
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
         LockType        =   2
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
      Begin VB.CommandButton cmdPerform 
         Appearance      =   0  'Flat
         Caption         =   "Perfor&mance"
         Height          =   280
         Left            =   10350
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.CommandButton cmdPosition 
         Appearance      =   0  'Flat
         Caption         =   "P&osition"
         Height          =   280
         Left            =   10350
         TabIndex        =   38
         Top             =   330
         Visible         =   0   'False
         Width           =   1250
      End
      Begin MSAdodcLib.Adodc Data3 
         Height          =   390
         Left            =   6120
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   688
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
         Caption         =   "HREMP"
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   12600
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportFileName  =   "c:\ihr\rgridsal.rpt"
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
   Begin Threed.SSPanel panEEDesc 
      Align           =   1  'Align Top
      Height          =   585
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   12840
      _Version        =   65536
      _ExtentX        =   22648
      _ExtentY        =   1032
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7560
         TabIndex        =   108
         Top             =   203
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   33
         Top             =   203
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1560
         TabIndex        =   32
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3120
         TabIndex        =   31
         Top             =   180
         Width           =   1740
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fesalary.frx":0000
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "Fesalary.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   9735
   End
End
Attribute VB_Name = "frmESALARY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEmptyNew
Dim UnionExecNone As Boolean

Dim orgSalary As Double
Dim orgSalary1 As Double
Dim OSalary, OSalCD, oJob, OEDate, ONDate, OReason, ONFSalary
Dim OPremium, OTOTAL, OvGroup, OVStep 'Vailtech
Dim oGrade
Dim Actn
Dim orgCurrent
Dim SavPAYP, OldPAYP, SavSalcd
Dim orgPosStDate As String
Dim dynaJobHIS As New ADODB.Recordset
Dim fglbJob$, fglbJobID&, fglbReason$
Dim fglbGrid$
Dim fglbPayrollID
Dim fglbSDate, fglbWhrs#, fglbBAND
Dim fglbPhrs, fglbDhrs
Dim fglbDiv
Dim OLambtonJob
Dim JobSnaps_PayScale(20) As Double '15 -> 20 Ticket #24983 Franks 01/31/2014
Dim JobSnaps_Salary_Code$
Dim JobSnaps_Salary_FTEHrs
Dim JobSnap_MidPoint!
Dim fSection As String

Dim fglbPCOld(4) As Double
Dim fglbAmtOld(6) As Currency
Dim fglbSHold@
Dim fglbGridType
Dim fglHredsem As String
Dim fglbNew As Boolean
Dim fglbFrmt As String
Dim flagFrmLoad As Boolean
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbJobList
Dim flgloaded As Boolean
Dim prompt As Boolean
Dim MailBody
Dim fglbNiagPhrs, fglbNiagWhrs
Dim locCountry As String, locCompDecHR
Dim xDefPosition As String
Dim xNP_VGroup(21)
Dim xNPVG_Cnt As Integer

Dim orgSalCD1  As String

Private Function AUDITSALY(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim rsTD As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim SQLQ As String, strFields As String
Dim xEffDateUpd, xSalUpd As Boolean

On Error GoTo AUDIT_ERR

AUDITSALY = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
'strFields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_GRID, AU_SALARY, AU_OLDSAL, AU_WHRS, AU_SALCD, "
'Added by Bryan 27/09/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'muskoka
    strFields = strFields & "AU_TOTAL, AU_VPREMIUM, AU_VGROUP, AU_VSTEP, "
End If
If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #22952 Franks 12/10/2012
     strFields = strFields & "AU_TOTAL, AU_VPREMIUM, "
End If
If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
    'North Perth Ticket #19209 Franks 05/18/2011
    'KN&V Ticket #21097 Franks 11/02/2011
    strFields = strFields & "AU_VGROUP, "
End If
strFields = strFields & "AU_JOB, AU_SEDATE, AU_SREASON, AU_PAYP, AU_OLDPAYP, "
strFields = strFields & "AU_SNDATE, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_JOB "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False
'~~~~~~~~~CHECK FOR NULL VALUES~~RAUBREY 6/19/97~~~~~~~~~~~
If IsNull(OSalary) Then
    OSalary = 0
End If

If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
    If IsNull(ONFSalary) Then
        ONFSalary = 0
    End If
End If

If IsNull(ONDate) Then
    ONDate = ""
Else
    If ONDate <> "01/01/01" Then    'THIS IS TO ENSURE THAT ONDATE
        ONDate = Trim(Str$(ONDate)) '    HAS NOT ALREADY BEEN SET
    End If                          '    TO A STRING IN Function
End If                              '    CurSHDate
If glbVadim And Not IsDate(ONDate) Then
    ONDate = "01/01/01"
End If

If IsNull(OEDate) Then
    OEDate = ""
Else
    If OEDate <> "01/01/01" Then 'THIS IS TO ENSURE THAT OEDATE HAS NOT ALREADY BEEN SET TO A STRING IN Function CurSHDate
        OEDate = Trim(Str$(OEDate))
    End If
End If

'do not know what should we do if there is salary changes

Dim xBatchID, UpdateAudit
Dim HRChanges As New Collection
Dim UptSalaryDate As Date

If fglbNew Or CVDate(Format(dlpDate(0), "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy")) Then
    UptSalaryDate = dlpDate(0)
Else
    UptSalaryDate = Date
End If

UpdateAudit = False
Dim HRSalary As New Collection
'Town of Aurora or City of Timmins
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Then 'Or glbCompSerial = "S/N - 2363W" Then
    'Or glbCompSerial = "S/N - 2276W" Then
    
    Dim LowSalary As New FieldInfo
    Dim LowSalCD As New FieldInfo
    Dim LowGrade As New FieldInfo
    Dim LowEDate As New FieldInfo
    Dim LowNDate As New FieldInfo
    Dim LowPAYP As New FieldInfo
    Dim LowReason As New FieldInfo
    Dim LowJob As New FieldInfo
    Dim LowDHRS As New FieldInfo
    Dim LowPayCat As New FieldInfo
    Dim rsOldSalary As New ADODB.Recordset
    
    SQLQ = "SELECT * from HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
    SQLQ = SQLQ & " AND SH_CURRENT <>0"
    
    If Not fglbNew Then
        SQLQ = SQLQ & " AND SH_ID<>" & Data1.Recordset("SH_ID")
    End If
    
    SQLQ = SQLQ & " ORDER BY SH_SALARY"
    rsOldSalary.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If chkCurrent = 0 Then
        'City of Timmins - do not pass anything if not current. Town of Aurora - removed (Ticket #20931)
        If glbCompSerial <> "S/N - 2375W" Then  'And glbCompSerial <> "S/N - 2378W" Then
            LowSalary.fdValue = 0: LowSalary.fdName = medsalary.DataField
        End If
    Else
        LowSalary.fdValue = Val(medsalary): LowSalary.fdName = medsalary.DataField
    End If
    
    LowSalCD.fdValue = lblSalCode: LowSalCD.fdName = lblSalCode.DataField
    LowGrade.fdValue = lblSalaryGrade: LowGrade.fdName = lblSalaryGrade.DataField
    
    If glbCompSerial <> "S/N - 2375W" Then   'City of Timmins
        LowEDate.fdValue = dlpDate(0): LowEDate.fdName = dlpDate(0).DataField
    End If
    
    LowNDate.fdValue = dlpDate(1): LowNDate.fdName = dlpDate(1).DataField
    LowPAYP.fdValue = clpCode(4): LowPAYP.fdName = clpCode(4).DataField
    LowReason.fdValue = clpCode(1): LowReason.fdName = clpCode(1).DataField
    LowJob.fdValue = clpPostCode: LowJob.fdName = "JH_JOB"
    LowDHRS.fdValue = 0: LowDHRS.fdName = "JH_DHRS"
    LowPayCat.fdValue = 0: LowPayCat.fdName = "JH_PAYROLL_CATEGORY"

    If Not rsOldSalary.EOF Then
        If rsOldSalary("SH_SALARY") < LowSalary.fdValue Or chkCurrent = 0 Then
            LowSalary.fdValue = rsOldSalary("SH_SALARY")
            LowSalCD.fdValue = rsOldSalary("SH_SALCD")
            LowGrade.fdValue = rsOldSalary("SH_GRADE")
            
            If glbCompSerial <> "S/N - 2375W" Then  'City of Timmins
                LowEDate.fdValue = rsOldSalary("SH_EDATE")
            End If
            
            LowNDate.fdValue = rsOldSalary("SH_NEXTDAT")
            LowPAYP.fdValue = rsOldSalary("SH_PAYP")
            LowReason.fdValue = rsOldSalary("SH_SREAS1")
            LowJob.fdValue = rsOldSalary("SH_JOB")
        End If
    End If
    
    'City of Timmins. Town of Aurora - removed (Ticket #20931)
    If (glbCompSerial = "S/N - 2375W") And chkCurrent = 0 Then  'Or glbCompSerial = "S/N - 2378W"
        'Do not pass anything if not current
    Else
        If isChanged_Salary(HRSalary, OSalary, LowSalary, True) Then UpdateAudit = True
    End If
    
    If isChanged_Salary(HRSalary, OSalCD, LowSalCD) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        'City of Niagara Falls has special logic to calculate the Hourly Rate so send
        'Hours per Day instead of Hours per Week
        If glbCompSerial = "S/N - 2276W" Then
            'To update HR_VADIM_SY_INTERFACE with Salary Amount
            txtVadPayRate.Text = ""
            txtVadSalRate.Text = ""
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbNiagPhrs, fglbDhrs, glbLEE_ID, txtPayrollID.Text)
            
        ElseIf glbCompSerial = "S/N - 2375W" Then  'City of Timmins
            rsTC.Open "SELECT JH_PHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_JOB='" & LowJob.fdValue & "' AND JH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
            If Not rsTC.EOF Then
                ''City of Timmins or Town of Aurora
                If (glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2378W") And chkCurrent = 0 Then
                    'Do not pass salary changes if not current
                Else
                    Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, rsTC("JH_PHRS"), rsTC("JH_WHRS"), glbLEE_ID, txtPayrollID.Text)
                End If
            End If
            rsTC.Close
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbPhrs, fglbWhrs, glbLEE_ID, txtPayrollID.Text)
        End If
        If isChanged_Field(HRChanges, oGrade, LowGrade, True) Then Debug.Print "" ' do nothing for the audit transfer
    End If
    
    
    If isChanged_Field(HRChanges, OEDate, LowEDate) Then UpdateAudit = True
    If isChanged_Field(HRChanges, ONDate, LowNDate) Then UpdateAudit = True
    If isChanged_Field(HRChanges, SavPAYP, LowPAYP) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OReason, LowReason) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oJob, LowJob) Then UpdateAudit = True
    
    'If OJOB <> LowJob.fdValue Then
        rsTC.Open "SELECT JH_DHRS,JH_PAYROLL_CATEGORY FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_JOB='" & LowJob.fdValue & "' AND JH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
        If Not rsTC.EOF Then
            LowDHRS.fdValue = rsTC("JH_DHRS")
            LowPayCat.fdValue = rsTC("JH_PAYROLL_CATEGORY")
        End If
        rsTC.Close
        If isChanged_Field(HRChanges, 0, LowDHRS) Then UpdateAudit = True
        If isChanged_Field(HRChanges, 0, LowPayCat) Then UpdateAudit = True
    'End If
    
    'Ticket #15070
    'Call Passing_Changes(HRChanges, Salary, "M", Date, glbLEE_ID, txtPayrollID.Text)
    Call Passing_Changes(HRChanges, Salary, "M", UptSalaryDate, glbLEE_ID, txtPayrollID.Text)

Else
    
    'City of Niagara Falls logic - Ticket #15542
    xEffDateUpd = False
    xSalUpd = False
    If fglbNew Then
        txtVadAddModDel.Text = "A"
    Else
        txtVadAddModDel.Text = "M"
    End If
    
    'Ticket #15542 - For Vadim only - Existing entry changed - delete the existing entries from Interface tables
    If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls
        If glbVadim And (Not fglbNew) And (CVDate(Format(dlpDate(0), "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy"))) Then
            'Call procedure to delete from SY_INTERFACE and SY_INTERFACE_BATCH tables
            Call Delete_Existing_Vadim("C")
            
            'Since the salary change and salary change related information got deleted in the above call
            'We have to add the salary change back again with the right changed information.
            'The change could be salary itself or effective date. If the Effective Date is changed then
            'it will be passed but not salary (if salary is not changed) and if salary is changed then it will be passed
            'but not related information (if these info has not changed).
            'Check what has changed
            If OSalary <> medsalary And OEDate <> dlpDate(0) Then
                'do nothing as it will be taken care below
            ElseIf OSalary <> medsalary Then
                'Pass Effective Date change info.
                'Force UpdateAudit to be true to pass the info to Vadim by reseting old value to blank
                If isChanged_Field(HRChanges, "", dlpDate(0)) Then UpdateAudit = True
                xEffDateUpd = True
            ElseIf OEDate <> dlpDate(0) Then
                'Pass Salary change info.
                'Force UpdateAudit to be true to pass the info to Vadim by reseting old value to blank
                If isChanged_Salary(HRSalary, "", medsalary, True) Then UpdateAudit = True
                If isChanged_Salary(HRSalary, "", lblSalCode) Then UpdateAudit = True
                xSalUpd = True
            End If
        ElseIf glbVadim And fglbNew Then
            'Update Employee History table of Vadim with End Date for the previous Salary record.
            If OSalary <> 0 And OEDate <> "" And Not IsNull(OEDate) Then
                'Comment enhacement - Ticket #16115
                'City of Niagara Falls - Ticket #15542
                'Update previous salary record in Vadim's HR_EMP_HIST table with End Date
                Call Update_VadimDB_HR_EMP_HISTORY(fglbPayrollID, OEDate, "", "", "", "M", DateAdd("d", -1, CVDate(dlpDate(0))))
            End If
        End If
    End If
    
    'Comment enhancement - Ticket #16115
    If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls  - Ticket #15542
        If xSalUpd = False Then
            If isChanged_Salary(HRSalary, OSalary, medsalary, True) Then UpdateAudit = True
            If glbCompSerial = "S/N - 2276W" And SavPAYP <> clpCode(4).Text Then   'City of Niagara Falls  'Ticket #16277/16276
                UpdateAudit = True
            End If
            If isChanged_Salary(HRSalary, OSalCD, lblSalCode) Then UpdateAudit = True
        End If
    Else
        If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
            If isChanged_Salary(HRSalary, OTOTAL, medTotal, True) Then UpdateAudit = True
        Else
            If isChanged_Salary(HRSalary, OSalary, medsalary, True) Then UpdateAudit = True
        End If
        If glbCompSerial = "S/N - 2276W" And SavPAYP <> clpCode(4).Text Then   'City of Niagara Falls  'Ticket #16277/16276
            UpdateAudit = True
        End If
        If isChanged_Salary(HRSalary, OSalCD, lblSalCode) Then UpdateAudit = True
    End If
        
    If glbVadim And UpdateAudit Then
        If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
            'To update HR_VADIM_SY_INTERFACE with Salary Amount
            txtVadPayRate.Text = ""
            txtVadSalRate.Text = ""
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbNiagPhrs, fglbDhrs, glbLEE_ID, fglbPayrollID) 'txtPayrollID.Text)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbPhrs, fglbWhrs, glbLEE_ID, fglbPayrollID) 'txtPayrollID.Text)
        End If
        'City of Kawartha Lakes - Pass Salary Grade to Probation Levels
        If glbCompSerial = "S/N - 2363W" Then
            If fglbNew Then
                If isChanged_Field(HRChanges, "", lblSalaryGrade, True) Then UpdateAudit = True
            Else
                If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then UpdateAudit = True
            End If
        Else
            
            'Ticket #24565 - District Municipality of Muskoka
            If glbCompSerial = "S/N - 2373W" Then
                'They want to transfer for 181W as well now - Nov 3rd 2014
                ''Ticket #24565 - if Union = '181W' then do not transfer Probation Date, Level and After Probation
                'If GetEmpData(glbLEE_ID, "ED_ORG") = "181W" Then
                '    'Do not transfer Probation Date, Level and After Probation
                'Else
                    If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then Debug.Print "" ' do nothing for the audit transfer
                'End If
            Else
                'Ticket #25412 - Town of Greater Napanee - No Probation Level, Date and After Probation Level to transfer
                'Ticket #25469 - City of Campbell River - do not transfer Probation levels
                If glbCompSerial <> "S/N - 2458W" And glbCompSerial <> "S/N - 2447W" Then
                    If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then Debug.Print "" ' do nothing for the audit transfer
                End If
            End If
        End If
    End If
    
    If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls  - Ticket #15542
        If xEffDateUpd = False Then
            If isChanged_Field(HRChanges, OEDate, dlpDate(0)) Then UpdateAudit = True
        End If
    Else
        If isChanged_Field(HRChanges, OEDate, dlpDate(0)) Then UpdateAudit = True
    End If
    
    If glbCompSerial <> "S/N - 2373W" Then 'DMuskoka - Ticket #24565 - Do not transfer Next Review Date
        If isChanged_Field(HRChanges, ONDate, dlpDate(1)) Then UpdateAudit = True
    End If
    
    If isChanged_Field(HRChanges, SavPAYP, clpCode(4)) Then UpdateAudit = True
    
    If isChanged_Field(HRChanges, OReason, clpCode(1)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oJob, clpPostCode) Then UpdateAudit = True
    
    If glbCompSerial = "S/N - 2373W" Then 'DMuskoka , ,  'Vailtech
        If isChanged_Field(HRChanges, OPremium, medPremium) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OTOTAL, medTotal) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OvGroup, txtVGroup) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OVStep, txtVStep) Then UpdateAudit = True
    End If
    
    If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
        'North Perth Ticket #19209 Franks 05/18/2011
        'KN&V Ticket #21097 Franks 11/02/2011
        If isChanged_Field(HRChanges, OvGroup, txtVGroup) Then UpdateAudit = True
    End If
    If glbCompSerial = "S/N - 2437W" Then 'Ticket #22952
        If isChanged_Field(HRChanges, OPremium, medPremium) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OTOTAL, medTotal) Then UpdateAudit = True
    End If
    
    'Ticket #15070
    'Call Passing_Changes(HRChanges, Salary, "M", Date, glbLEE_ID, txtPayrollID.Text)
    Call Passing_Changes(HRChanges, Salary, "M", UptSalaryDate, glbLEE_ID, txtPayrollID.Text)
    
    'City of Niagara Falls - Ticket #15542
    If glbVadim And glbCompSerial = "S/N - 2276W" Then
        If txtVadAddModDel.Text = "A" Then
            'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
            Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, UptSalaryDate, "", Val(lblSalaryGrade), clpPostCode, "A")
        ElseIf txtVadAddModDel.Text = "M" Then
            'Modify the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
            If Val(oGrade) <> Val(lblSalaryGrade) Then
                Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, UptSalaryDate, Val(oGrade), Val(lblSalaryGrade), clpPostCode, "M")
            End If
            If OEDate <> dlpDate(0) Then
                'if the Effective Date changes then delete the original record and add a complete
                'new record with the new Effective Date. This is to avoid confusion on when
                'the change should be effective.
                'Delete the salary record from Vadim's HR_EMP_HIST table storing the history of Rate changes
                Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID.Text, OEDate, Val(oGrade), "", clpPostCode, "D")
                
                'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
                Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, UptSalaryDate, "", Val(lblSalaryGrade), clpPostCode, "A")
                'Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, UptSalaryDate, IIf(Not IsNull(OEDate), Format(OEDate, "YYYY/MM/DD"), OEDate), Format(dlpDate(0), "YYYY/MM/DD"), clpPostCode, "M")
            End If
            If oJob <> clpPostCode Then
                Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, UptSalaryDate, oJob, clpPostCode, clpPostCode, "M")
            End If
        End If
    End If
End If
If UpdateAudit Then GoTo MODUPD Else GoTo MODNOUPD


MODUPD:
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_GRID") = clpGrid.Text


If Trim(Str$(OSalary)) <> medsalary Or SavSalcd <> lblSalCode Then 'Trim(Str$ added by RAUBREY 6/3/97
    If glbFrench Then
        rsTA("AU_SALARY") = Replace(medsalary, ",", ".")
        rsTA("AU_OLDSAL") = Replace(OSalary, ",", ".")
    Else
        rsTA("AU_SALARY") = medsalary
        rsTA("AU_OLDSAL") = OSalary
    End If
    If Len(lblWhrs) > 0 Then
        rsTA("AU_WHRS") = lblWhrs
    End If
    rsTA("AU_SALCD") = lblSalCode 'laura febr 2, 1998
    'Ticket #20333 Franks 05/16/2011, Jerry asked always pass JOB for a new salary
    rsTA("AU_JOB") = clpPostCode.Text
End If

'Added by Bryan 27/09/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'muskoka
    'If (OPremium) <> medPremium Or (OTotal) <> medTotal Or OvGroup <> txtVGroup Or OVStep <> txtVStep Then
        rsTA("AU_TOTAL") = medTotal
        rsTA("AU_VPREMIUM") = IIf(medPremium = "", 0, medPremium)
        rsTA("AU_VGROUP") = cboVGRoup 'txtVGroup
        rsTA("AU_VSTEP") = cboVStep 'txtVStep
    'End If
End If
If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
    'North Perth Ticket #19209 Franks 05/18/2011
    'KN&V Ticket #21097 Franks 11/02/2011
    rsTA("AU_VGROUP") = txtVGroup.Text
    rsTA("AU_JOB") = clpPostCode.Text
End If
If glbCompSerial = "S/N - 2437W" Then 'Ticket #22952
        If isChanged_Field(HRChanges, OTOTAL, medTotal) Then rsTA("AU_TOTAL") = medTotal
        If isChanged_Field(HRChanges, OPremium, medPremium) Then rsTA("AU_VPREMIUM") = IIf(medPremium = "", 0, medPremium)
End If

If glbInsync Then
    rsTA("AU_JOB") = clpPostCode.Text
    rsTA("AU_SEDATE") = dlpDate(0).Text
    rsTA("AU_SREASON") = clpCode(1).Text
Else
    If oJob <> clpPostCode.Text Then rsTA("AU_JOB") = clpPostCode.Text
    If OEDate <> dlpDate(0).Text Then rsTA("AU_SEDATE") = dlpDate(0).Text
    If ACTX = "A" Then
        rsTA("AU_SREASON") = clpCode(1).Text
    Else
        If OReason <> clpCode(1).Text Then rsTA("AU_SREASON") = clpCode(1).Text
    End If
End If

If SavPAYP <> clpCode(4).Text Then
    If Len(clpCode(4).Text) > 0 Then
        rsTA("AU_PAYP") = clpCode(4).Text
    Else
        rsTA("AU_PAYP") = "-"
    End If
    If Not IsNull(SavPAYP) Then
        If SavPAYP <> "" Then rsTA("AU_OLDPAYP") = SavPAYP
    End If
Else
    'If Val(clpCode(4).Text) = 0 Then
    'Ticket #23736 Franks 05/10/2013
    If IsNumeric(clpCode(4).Text) And Val(clpCode(4).Text) = 0 Then
        rsTA("AU_PAYP") = Null
    Else
        'rsTA("AU_PAYP") = Val(clpCode(4).Text)
        'Ticket #21152 Franks 11/15/2012 can't change it to numeric, it drop 0 from "0019", also PP can be string, e.g. "BW"
        rsTA("AU_PAYP") = (clpCode(4).Text)
    End If
    'If SavPAYP <> "" Then
    If Not SavPAYP = clpCode(4).Text Then
         rsTA("AU_OLDPAYP") = SavPAYP ' Val(SavPAYP) clpCode(4).Text
    End If
End If


If IsDate(dlpDate(1).Text) Then                   '13Aug99 js
    If ONDate <> dlpDate(1).Text Then             '
        rsTA("AU_SNDATE") = dlpDate(1).Text      '
    End If                                  '
Else                                        '
    rsTA("AU_SNDATE") = Null                '
End If                                      '

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID


If glbCompSerial = "S/N - 2290W" Then
    rsTA("AU_LDATE") = Date
Else
    If Actn = "A" Then
        If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
            rsTA("AU_LDATE") = Format(DateAdd("d", 14, dlpDate(0)), "SHORT DATE")
        Else
            'rsTA("AU_LDATE") = dlpDate(0).Text
            If CVDate(dlpDate(0).Text) > CDate(Date) Then
                rsTA("AU_LDATE") = CVDate(dlpDate(0).Text)
            Else
                rsTA("AU_LDATE") = Date
            End If
        End If
    Else
        If CVDate(dlpDate(0).Text) > CDate(Date) Then
            rsTA("AU_LDATE") = CVDate(dlpDate(0).Text)
        Else
            rsTA("AU_LDATE") = Date
        End If
    End If
End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
If glbMulti Then
    rsTA("AU_PAYROLL_ID") = txtPayrollID
Else
    Dim rsEmp As New ADODB.Recordset
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    If OSalary <> medsalary Then
        rsTA("AU_JOB") = clpPostCode.Text '# 7644
    End If
End If
rsTA.Update
' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
Call Pause(0.5)


'~~~~~~~~~~~~~~~~~~~~~~~~

MODNOUPD:
AUDITSALY = True

Exit Function
AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '28July99 js
Resume Next
End Function
Private Sub TermRehireAudit(rsTA As ADODB.Recordset)
Dim rsTC As New ADODB.Recordset
Dim SQLQ, xTilPayID
    rsTC.Open "SELECT * FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If rsTC.EOF Then Exit Sub
    'If IsNull(rsTC("ED_PAYROLL_ID")) Then Exit Sub
    'Termination Data
    If Not glbCompSerial = "S/N - 2369W" Then    'TS Tech
        rsTA.AddNew
        rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
        rsTA("AU_NEWEMP") = "N"
        rsTA("AU_SURNAME") = rsTC("ED_SURNAME") '
        rsTA("AU_FNAME") = rsTC("ED_FNAME")
        rsTA("AU_DOT") = glbChgTermDate
        rsTA("AU_TREAS") = glbChgTermReason
        rsTA("AU_COMPNO") = "001"
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
        rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
        rsTA("AU_PAYP") = OldPAYP
        rsTA("AU_DIVUPL") = rsTC("ED_DIV")
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_VERSION") = "ADPTRA" 'Ticket# 7768
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "T"
        rsTA.Update
    End If
    'New Hire Data
    xTilPayID = ""
    If Not IsNull(rsTC("ED_PAYROLL_ID")) Then
        xTilPayID = rsTC("ED_PAYROLL_ID")
    End If
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_ADMINBY_TABL") = "EDAB": rsTA("AU_LANG1_TABL") = "EDL1":: rsTA("AU_LANG2_TABL") = "EDL1"
    rsTA("AU_DIV") = rsTC("ED_DIV")
    rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
    rsTA("AU_TITLE") = rsTC("ED_TITLE")
    rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
    rsTA("AU_FNAME") = rsTC("ED_FNAME")
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
    rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
    rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
    rsTA("AU_CITY") = rsTC("ED_CITY")
    rsTA("AU_PROV") = rsTC("ED_PROV")
    rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
    rsTA("AU_PCODE") = rsTC("ED_PCODE")
    rsTA("AU_PHONE") = rsTC("ED_PHONE")
    rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
    rsTA("AU_DIVUPL") = rsTC("ED_DIV")
    rsTA("AU_SEX") = rsTC("ED_SEX")
    rsTA("AU_SMOKER") = IIf(rsTC("ED_SMOKER"), "Yes", "No")
    rsTA("AU_DOB") = rsTC("ED_DOB")
    rsTA("AU_SIN") = rsTC("ED_SIN")
    rsTA("AU_DEPT_GL") = rsTC("ED_GLNO")
    rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
    rsTA("AU_NEWEMP") = "Y"
    rsTA("AU_PTUPL") = rsTC("ED_PT")
    rsTA("AU_LOC") = rsTC("ED_LOC")
    rsTA("AU_TD1") = rsTC("ED_TD1")
    rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
    rsTA("AU_PROVFORM") = rsTC("ED_PROVFORM")
    rsTA("AU_PROVAMT") = rsTC("ED_PROVAMT")
    rsTA("AU_OLDTD1") = 0
    rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
    rsTA("AU_REGION") = rsTC("ED_REGION")
    rsTA("AU_SECTION") = rsTC("ED_SECTION")
    rsTA("AU_HOMEOPRTNBR") = rsTC("ED_HOMEOPRTNBR")
    rsTA("AU_HOMELINE") = rsTC("ED_HOMELINE")
    rsTA("AU_HOMESHIFT") = rsTC("ED_HOMESHIFT")
    rsTA("AU_HOMEWRKCNT") = rsTC("ED_HOMEWRKCNT")
    rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
    rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
    rsTA("AU_SSN") = rsTC("ED_SSN")
 
    rsTA("AU_DEPTEDATE") = rsTC("ED_DEPTEDATE")
    rsTA("AU_DIVEDATE") = rsTC("ED_DIVEDATE")
    rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
    rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
    rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
    rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
    rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
    rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")
    rsTA("AU_BADGEID") = rsTC("ED_BADGEID")
    rsTA("AU_MIDNAME") = rsTC("ED_MIDNAME")
    rsTA("AU_ALIAS") = rsTC("ED_ALIAS")
    'Employee Status
    rsTA("AU_EMP") = rsTC("ED_EMP") 'clpCode(1) '
    rsTA("AU_LDATE") = Format(Now, "SHORT DATE")
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    rsTA.Update
    
    '------BANK Information Begin
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

    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_DIVUPL") = rsTC("ED_DIV")
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    If Len(xTilPayID) > 0 Then rsTA("AU_Payroll_ID") = xTilPayID
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
    End If
    rsTC.Close
    rsTC.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenStatic
    If Not rsTC.EOF Then
        rsTA("AU_SALARY") = rsTC("SH_SALARY")
        rsTA("AU_WHRS") = rsTC("SH_WHRS")
        rsTA("AU_SALCD") = rsTC("SH_SALCD")
        rsTA("AU_SEDATE") = rsTC("SH_NEXTDAT")
        If glbCompSerial = "S/N - 2369W" Then    'TS Tech
            rsTA("AU_PAYP") = clpCode(4).Text
        Else
            rsTA("AU_PAYP") = rsTC("SH_PAYP")
        End If
    End If
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "A"
    If Len(xTilPayID) > 0 Then rsTA("AU_Payroll_ID") = xTilPayID
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
        rsTA("AU_EMPNBR") = glbLEE_ID
        rsTA("AU_LDATE") = Date
        rsTA("AU_LUSER") = glbUserID
        rsTA("AU_LTIME") = Time$
        rsTA("AU_UPLOAD") = "N"
        rsTA("AU_TYPE") = "A"
        If Len(xTilPayID) > 0 Then rsTA("AU_Payroll_ID") = xTilPayID
        rsTA.Update
        rsTC.MoveNext
    Loop
    rsTC.Close
    '------Other Earnings End

End Sub



Private Sub cboVGRoup_Click()
If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
    'North Perth Ticket #19209 Franks 05/18/2011
    'KN&V Ticket #21097 Franks 11/02/2011
    txtVGroup.Text = getVGrpTxt(cboVGRoup.Text)
Else
    txtVGroup.Text = cboVGRoup.Text
End If
End Sub
Private Function getVGrpcno(xText)
Dim retVal
If glbCompSerial = "S/N - 2429W" Then 'Municipality of North Perth
    retVal = -1
    'Ticket #21232 Franks 01/26/2012 - begin
    'Select Case xText
    'Case "H"
    '    retVal = 0
    'Case "H2"
    '    retVal = 1
    'Case "H3"
    '    retVal = 2
    'Case "H4"
    '    retVal = 3
    'Case "S"
    '    retVal = 4
    'End Select
    retVal = getNPIndex(xText)
    'Ticket #21232 Franks 01/26/2012 - end
    getVGrpcno = retVal
End If
If glbCompSerial = "S/N - 2437W" Then
    retVal = -1
    Select Case xText
    Case "SB"
        retVal = 0
    Case "HB"
        retVal = 1
    Case "H2"
        retVal = 2
    End Select
    getVGrpcno = retVal
End If
End Function
Private Function getNPIndex(xText)
Dim I As Integer
Dim xTemp As String
Dim retVal
    retVal = -1
    xTemp = Left(xText & "  ", 3) & "-"
    For I = 0 To xNPVG_Cnt ' 21
        'If InStr(1, xNP_VGroup(I), xTemp) > 0 Then
        If Left(xNP_VGroup(I), 4) = xTemp Then
            retVal = I
        End If
    Next
    getNPIndex = retVal
End Function
Private Function getVGrpTxt(xText)
Dim retVal As String
If glbCompSerial = "S/N - 2429W" Then 'Municipality of North Perth
    retVal = ""
    'Ticket #21232 Franks 01/26/2012 - begin
    'Select Case xText
    'Case "HOURLY"
    '    retVal = "H"
    'Case "HOURLY2"
    '    retVal = "H2"
    'Case "HOURLY3"
    '    retVal = "H3"
    'Case "HOURLY4"
    '    retVal = "H4"
    'Case "SALARY"
    '    retVal = "S"
    'End Select
    retVal = Trim(Left(xText, 3))
    'Ticket #21232 Franks 01/26/2012 - end
    getVGrpTxt = retVal
End If
If glbCompSerial = "S/N - 2437W" Then
    retVal = ""
    Select Case xText
    Case "SB"
        retVal = "SB"
    Case "HB"
        retVal = "HB"
    Case "H2"
        retVal = "H2"
    End Select
    getVGrpTxt = retVal
End If
End Function

Private Sub cboVStep_Click()
txtVStep = cboVStep.Text
End Sub

Private Sub chkCurrent_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function Chkpos()
Dim SQLQ As String, Msg$
Dim xPosFind

On Error GoTo ChkPos_Err

Chkpos = False

If Len(dlpPosStDate.Text) < 1 Then
    ' If pos. start date is missing in multi, it means they didn't enter a valid position
    If glbMulti Then
        MsgBox "Position does not exist in Position History file.  Please correct this before continuing.", vbOKOnly + vbExclamation, "Position Not Found"
         clpPostCode.SetFocus
    Else
        Msg$ = "Position Start Date is required"
        dlpPosStDate.SetFocus
        MsgBox Msg$
    End If
    Exit Function
Else
    If Not IsDate(dlpPosStDate.Text) Then
        Msg$ = "Not a Valid Position Start Date"
        dlpPosStDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If

End If

If Len(dlpDate(0).Text) < 1 Then
    Msg$ = "Effective Date is required"
    dlpDate(0).SetFocus
    MsgBox Msg$
    Exit Function
End If

If Len(clpPostCode.Text) > 0 Then
    If clpPostCode.Caption = "Unassigned" Then
        MsgBox "Position Code is invalid"
         clpPostCode.SetFocus
        Exit Function
    End If
Else
    If clpPostCode.Caption = "Unassigned" Then
        MsgBox "Position Code is required"
         clpPostCode.SetFocus
        Exit Function
    End If
End If
xPosFind = False
If Not Set_Position(clpPostCode.Text, False) Then
    Msg$ = "No position <" & clpPostCode.Text & "> found "
    Msg$ = Msg$ & Chr(10) & "Please review positions from Position History!"
    MsgBox Msg$
    Exit Function
End If
If dlpPosStDate.Text <> fglbSDate Then
    MsgBox "Start Date in the Salary History is different than the Position History!"
End If
If glbMultiGrid Then
    If Len(clpGrid.Text) <= 0 Then
        MsgBox lStr("Grid Category is required")
        clpGrid.SetFocus
        Exit Function
    Else
        If clpGrid.Caption = "Unassigned" Then
            MsgBox lStr("Grid Category is required")
            clpGrid.SetFocus
            Exit Function
        End If
    End If
End If
If glbMulti And glbVadim Then
    Dim rsChkJob As New ADODB.Recordset
    If chkCurrent Then
        rsChkJob.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID & " AND JH_PAYROLL_ID='" & txtPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
        If rsChkJob.EOF Then
            Msg$ = "No Payroll ID found in the Current Positions"
            Msg$ = Msg$ & Chr(10) & "Please review positions from Position History!"
            MsgBox Msg$
            txtPayrollID.SetFocus
            Exit Function
        End If
        rsChkJob.Close
    End If
End If

If glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
    'North Perth Ticket #19209 Franks 05/18/2011
    'KN&V Ticket #21097 Franks 11/02/2011
    If Len(cboVGRoup.Text) = 0 Then
        MsgBox "Pay Type is required."
        Exit Function
    End If
End If

Chkpos = True

Exit Function

ChkPos_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdChPos", "HR_JOB_HISTORY", "Change Position")
Resume Next

End Function

Private Function chkSalHist()
Dim X%
Dim SQLQ As String, Msg$, dd&
Dim DgDef As Variant, Title$, Response%, DCurSHDate  As Variant
Dim rsEmp As New ADODB.Recordset
Dim dtEmpDOH As Date
chkSalHist = False

On Error GoTo chkSalH_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Reason Code is required"
    clpCode(1).SetFocus
    Exit Function
Else
    For X% = 1 To 4
        If X% < 4 Then
            If Len(clpCode(X%).Text) = 0 Then
                medPercentChng(X%) = 0
                medAmtChng(X%) = 0
            End If
        End If
        If clpCode(X%).Caption = "Unassigned" Then
            If X% < 4 Then
                MsgBox "Reason Code must be valid"
            Else
                MsgBox "Pay Period Code must be valid"
            End If
            clpCode(X%).SetFocus
            Exit Function
        End If
    Next X%
End If
If glbVadim Then
    If glbMulti Then 'Ticket# 7751
        If Len(txtPayrollID.Text) = 0 Then
            MsgBox "Payroll ID is required"
            txtPayrollID.SetFocus
            Exit Function
        End If
    End If
End If

If glbPayWeb Or glbVadim Or glbLambton Or glbInsync Or glbCompSerial = "S/N - 2348W" _
Or glbCompSerial = "S/N - 2351W" Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2370W" _
Or glbCompSerial = "S/N - 2380W" _
Or (glbWFC) Then ' And fSection = "GREN") Then
    If Len(clpCode(4).Text) = 0 Then
        If Not glbCompSerial = "S/N - 2386W" And Not glbCompSerial = "S/N - 2382W" Then     'The Walter Fedy Partnership Ticket #14003
            'Samuel
            MsgBox lStr("Pay Period Code is required")
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
    
End If
' -----
If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2425W" Or glbCompSerial = "S/N - 2383W" Then
'Ticket #16478 Samuel, Ticket #18221 - Four Villages CHC
'2383 - Town of Orangeville Ticket #18844 Franks 01/13/2011
    If Len(clpCode(4).Text) = 0 Then
        MsgBox "Pay Type is required"
        clpCode(4).SetFocus
        Exit Function
    End If
End If
If Len(medsalary) < 1 Then
    If fraSalary.Enabled = True Then medsalary.SetFocus
    MsgBox "Salary is required"
    If medsalary.Enabled Then medsalary.SetFocus
    Exit Function
End If
If medsalary <= 0 Then
    If fraSalary.Enabled = True Then medsalary.SetFocus
    MsgBox "Salary is required"
    If medsalary.Enabled Then medsalary.SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
    If Not IsNumeric(medNFacSalary) And Len(Trim(medNFacSalary)) > 0 Then
        MsgBox "Invalid Non-Factored Salary"
        medNFacSalary.SetFocus
        Exit Function
    End If
End If

' -----
'Hemu - 06/18/2003 Begin - Incase the 'Per' has no value
    If comPayPer.Text = "" Then
        MsgBox "Per cannot be blank"
        comPayPer.SetFocus
        Exit Function
    End If

'Hemu - 06/18/2003 End

If glbWFC Then 'Frank 09/24/04 Ticket# 6962
    If clpCode(0).Visible And Len(clpCode(0).Text) < 1 Then
        Msg$ = "Plant is required"
        clpCode(0).SetFocus
        MsgBox Msg$
        Exit Function
    End If
    If txtFiscalYear.Visible And Len(txtFiscalYear) < 1 Then
        Msg$ = "Fiscal Year is required"
        txtFiscalYear.SetFocus
        MsgBox Msg$
        Exit Function
    End If
    If cmbMarketLine.Visible And Len(cmbMarketLine.Text) < 1 Then
        Msg$ = "Market Line is required"
        cmbMarketLine.SetFocus
        MsgBox Msg$
        Exit Function
    End If
    If Trim(comPayPer.Text) = "Annum" Or Trim(comPayPer.Text) = "Monthly" Then
        If Not IsDate(dlpDate(1).Text) Then
            Msg$ = "Next Review is required"
            dlpDate(1).SetFocus
            MsgBox Msg$
            Exit Function
        End If
    End If
End If

'Ticket #24482 - Town of Caledon - Using the VGroup field to store the Job's Division to create uniqueness between
'multiple same Position and Start Date positions linked to Salary.
If glbCompSerial = "S/N - 2182W" Then
    If clpDiv.Caption = "Unassigned" And Len(clpDiv.Text) > 0 Then
        MsgBox lStr("Division Code must be valid")
        clpDiv.SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(0).Text) < 1 Then
    Msg$ = "Effective Date is required"
    dlpDate(0).SetFocus
    MsgBox Msg$
    Exit Function
Else
    If Not IsDate(dlpDate(0).Text) Then
        Msg$ = "Not a Valid Effective Date"
        dlpDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    Else
        DCurSHDate = CurSHDate()
        If DCurSHDate > 0 Then    ' 0 if no current record out there
           DCurSHDate = CVDate(DCurSHDate)
           If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) <> 0 Then
                Call ChangeEDateAudit(DCurSHDate)
                
           End If
        End If
        If glbSetSal Then
            If DCurSHDate > 0 Then    ' 0 if no current record out there
                DCurSHDate = CVDate(DCurSHDate)
                If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) <= 0 And chkCurrent Then
                    Msg$ = "Warning...you cannot add or edit a record with a date"
                    'Msg$ = Msg$ & Chr(10) & "the same or later than your most current record."
                    Msg$ = Msg$ & " the same or later than your most current record."
                    Msg$ = Msg$ & Chr(10) & "If you need to edit current salary, "
                    'Msg$ = Msg$ & Chr(10) & "go to Salary screen under Employee Menu."
                    Msg$ = Msg$ & "go to Salary screen under Employee menu \ Work History/Compensation."
                    MsgBox Msg$, vbExclamation
                    dlpDate(0).SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If
'Hemu 05/13/2003 Begin - Effective Date and Original Hire Date
If Len(dlpDate(0).Text) > 0 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Effective Date is not a valid date"
        dlpDate(0).SetFocus
        Exit Function
    End If
    If Not glbLambton Then
        rsEmp.Open "SELECT ED_DOH,ED_SENDTE FROM HREMP WHERE ED_EMPNBR = " & lblEENum, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
            If glbSamuel Then 'Ticket #21202 Franks 11/16/2011
                If Not IsNull(rsEmp("ED_SENDTE")) Then
                    dtEmpDOH = rsEmp("ED_SENDTE")
                    If DaysBetween(rsEmp("ED_SENDTE"), dlpDate(0).Text) < 0 Then
                        MsgBox "Effective Date can not be prior to " & lStr("Seniority") & ""
                        dlpDate(0).SetFocus
                        rsEmp.Close
                        Exit Function
                    End If
                End If
            Else
                If rsEmp("ED_DOH") <> "" Then
                    dtEmpDOH = rsEmp("ED_DOH")
                    If DaysBetween(rsEmp("ED_DOH"), dlpDate(0).Text) < 0 Then
                        MsgBox "Effective Date can not be prior to Original Hire date"
                        dlpDate(0).SetFocus
                        rsEmp.Close
                        Exit Function
                    End If
                End If
            End If
        End If
        rsEmp.Close
    End If
End If
'Hemu 05/13/2003 End

DCurSHDate = CurSHDate()
If Not fglbNew And glbMediPay Then
    Dim OtherChange
    If SavPAYP <> clpCode(4) Then
        OtherChange = False

        If CDbl(OSalary) <> CDbl(medsalary) Then OtherChange = True
        If OSalCD <> lblSalCode Then OtherChange = True
        If OEDate <> dlpDate(0) Then OtherChange = True
        If ONDate <> dlpDate(1) Then OtherChange = True
        If OReason <> clpCode(1) Then OtherChange = True
        If oJob <> clpPostCode Then OtherChange = True
        If OtherChange Then
            Msg$ = "Warning, you can not change Salary information with the Client # transfer."
            Msg$ = Msg$ & Chr(10) & "Please cancel the changes."
            DgDef = MB_OK + MB_ICONQUESTION + MB_DEFBUTTON2
            Response% = MsgBox(Msg$) ', DgDef, "Warning!")
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
End If
If glbAddHisWarning And Actn = "A" And (Not glbSetSal) Then
    If DCurSHDate > 0 Then    ' 0 if no current record out there
        DCurSHDate = CVDate(DCurSHDate)
        If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) > 0 Then
            Msg$ = "Warning, you can not add a record with a date"
            Msg$ = Msg$ & Chr(10) & "earlier than your most current record."
            DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
            Response% = MsgBox(Msg$) ', DgDef, "Warning!")
            dlpDate(0).SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #24565 - They want this to be only mandatory when New Hire
If NewHireForms.count > 0 Then
    'Ticket #19113
    If glbCompSerial = "S/N - 2373W" Then 'Dist. of Muskoka
        If Not IsDate(dlpDate(1).Text) Then
            Msg$ = "Next Review is required"
            dlpDate(1).SetFocus
            MsgBox Msg$
            Exit Function
        End If
    End If
End If

If Len(dlpDate(1).Text) > 0 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Next Review Date is invalid"
        dlpDate(1).SetFocus
        Exit Function
    End If
        'Hemu - 05/13/2003 Begin
    If DaysBetween(dtEmpDOH, dlpDate(1).Text) < 0 Then
        MsgBox "Next Review date can not be prior to Original Hire date"
        dlpDate(1).SetFocus
        Exit Function
    End If
    'Hemu - 05/13/2003 End
    dd& = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))
    If dd& < 0 Then
        Msg$ = "Next Review precedes Effective date of salary "
        dlpDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    End If
Else
    If glbLinamar Then '(chkCurrent Or Actn = "A") Then 'Ticket #15546
        MsgBox "Next Review Date is required"
        dlpDate(1).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(2).Text) > 0 Then
    If Not IsDate(dlpDate(2).Text) Then
        MsgBox "Transaction Date is invalid"
        dlpDate(2).SetFocus
        Exit Function
    End If
End If


' dkostka - 03/20/2002 - Added check for user compa box
' if it's woodbridge and compa is set by user, they have to enter a value.
'Ticket# 6962 don't need this any more 09/24/04 Frank
'If glbWFC And optUserSys(1).Value = True And (mskCampa.Text = "" Or Not IsNumeric(mskCampa.Text)) Then
'    MsgBox "If Compa Ratio is set to user, a value must be entered.", vbExclamation + vbOKOnly, "Value Required"
'    mskCampa.SetFocus
'    Exit Function
'End If



'Frank 08/27/03 - Pay Period is mandatory for Soroc
If glbSoroc Or glbSyndesis Or glbCompSerial = "S/N - 2229W" Or glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2171W" Then  'Soroc, Syndesis,Inscape,'Ticket #24504 SPC 'Ticket #26971 Russell A
    If Len(clpCode(4).Text) < 1 Then
        Msg$ = lStr("Pay Period is required")
        clpCode(4).SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2390W" Then  'Collectcorp Ticket #14889
    If glbEmpCountry = "U.S.A." Then
        If Len(clpCode(4).Text) < 1 Then
            Msg$ = lStr("Pay Period is required for the U.S.A. employees.")
            clpCode(4).SetFocus
            MsgBox Msg$
            Exit Function
        End If
    End If
End If

If (glbCompSerial = "S/N - 2242W") Then  'C.C.A.C. London & Middlesex - Ticket #6718
    If Len(clpCode(4).Text) = 0 Then
        MsgBox "Client # is required"
        clpCode(4).SetFocus
        Exit Function
    End If
    
    If Not clpCode(4).ListChecker Then
        MsgBox "Client # must be valid"
        clpCode(4).SetFocus
    End If
End If
If (glbCompSerial = "S/N - 2387W") Then  'Bird Packaging Limited  - TTicket #13166
    If Len(clpCode(4).Text) = 0 Then
        MsgBox "Company Code is required"
        clpCode(4).SetFocus
        Exit Function
    End If
    
    If Not clpCode(4).ListChecker Then
        MsgBox "Company Code must be valid"
        clpCode(4).SetFocus
    End If
End If

If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #21097 Franks 11/02/2011
    If Len(cboVGRoup.Text) = 0 Then
        MsgBox ("Pay Type is required")
        cboVGRoup.SetFocus
        Exit Function
    End If
    'KN&V Ticket #22952 Franks 12/10/2012
    If Len(medPremium.Text) = 0 Then
        MsgBox (lblTitle(16).Caption & " is required")
        medPremium.SetFocus
        Exit Function
    Else
        If Not IsNumeric(medPremium.Text) Then
            MsgBox (lblTitle(16).Caption & " is not numeric")
            medPremium.SetFocus
            Exit Function
        End If
    End If
    If Len(medTotal.Text) = 0 Then
        MsgBox (lblTitle(18).Caption & " is required")
        medTotal.SetFocus
        Exit Function
    Else
        If Not IsNumeric(medTotal.Text) Then
            MsgBox (lblTitle(18).Caption & " is not numeric")
            medTotal.SetFocus
            Exit Function
        End If
    End If
End If

If (glbCompSerial = "S/N - 2409W") Then 'Ticket #30066 Franks - Skylark Children
    If Len(clpCode(4).Text) = 0 Then
        MsgBox lStr("Pay Period Code is required")
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2429W" Then 'North Perth Ticket #19209 Franks 05/19/2011
    If Len(clpCode(4).Text) = 0 Then
        MsgBox lStr("Pay Period Code is required")
        clpCode(4).SetFocus
        Exit Function
    End If
    If Len(cboVGRoup.Text) = 0 Then
        MsgBox ("Pay Type is required")
        cboVGRoup.SetFocus
        Exit Function
    End If
    If chkCurrent.Value Then 'current only
        If clpPostCode.Text = xDefPosition Then
            ''If clpCode(4).Text = "4" Or clpCode(4).Text = "6" Then
            ''    If Not cboVGRoup.Text = "SALARY" Then
            ''        MsgBox ("Pay Type must be 'SALARY' for Default Position if the " & lStr("Pay Period") & " is '4' or '6'")
            ''        cboVGRoup.SetFocus
            ''        Exit Function
            ''    End If
            ''End If
            ''If Not (clpCode(4).Text = "4" Or clpCode(4).Text = "6") Then
            ''    If Not cboVGRoup.Text = "HOURLY" Then
            ''        MsgBox ("Pay Type must be 'HOURLY' for Default Position if the " & lStr("Pay Period") & " is not '4' and '6'")
            ''        cboVGRoup.SetFocus
            ''        Exit Function
            ''    End If
            ''End If
        End If
        'check duplicate
        'Ticket #24158 - No checking for Municipality of North Perth
        If glbCompSerial <> "S/N - 2429W" Then
            If chkDupVGroup(txtVGroup.Text, fglbNew) Then
                MsgBox "Duplicated Pay Type " & cboVGRoup.Text & " found in Current Salary Records."
                cboVGRoup.SetFocus
                Exit Function
            End If
        End If
    End If
End If

If DCurSHDate = 0 Then DCurSHDate = dlpDate(0).Text   'New Record
If IsDate(DCurSHDate) Then
    If DateDiff("d", CVDate(dlpDate(0).Text), DCurSHDate) <= 0 Then
        If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Then  'Town of Aurora and Timmmins
            If Not AUDITSALY(Actn) Then MsgBox "ERROR - AUDIT FILE"
        ElseIf Not glbMulti Or chkCurrent = True Then
            If Not AUDITSALY(Actn) Then MsgBox "ERROR - AUDIT FILE"
        End If
    End If
End If

If glbCompSerial = "S/N - 2259W" Then 'Oxford 'Ticket #21599 Franks 03/01/2012
    If glbMulti Then
        If fglbNew Then 'new record
            If CheckDuplCurrent(glbLEE_ID, clpPostCode.Text) Then
                Msg$ = "There is another current Salary for the same Position Code '" & clpPostCode.Text & "' " & Chr(10)
                Msg$ = Msg$ & "You can't have two current Salaries for the same Code" & Chr(10)
                Msg$ = Msg$ & "Please uncheck the Current Salary Record flag for the previous Current Salary." & Chr(10)
                MsgBox Msg$
                Exit Function
            End If
        End If
    End If
End If


'Ticket #21791 Franks 04/09/2012
If glbSamuel Then
    If UCase(clpCode(4).Text) = "V" Then
        If comPayPer.Text = "Annum" Then
            MsgBox "Please ensure salary entered is an hourly rate if Pay Type is 'V' "
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2443W" Then 'Walters Inc - Ticket #23231 Franks 02/16/2013
    If Not glbtermopen Then
        If GetEmployeeInfo("ED_DEPTNO") = "02" Then
            If Not comPayPer.Text = "Annum" Then
                MsgBox "Per must be 'Annum' if Department is '02'"
                comPayPer.SetFocus
                Exit Function
            End If
        End If
    End If
End If

chkSalHist = True

Exit Function

chkSalH_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSal", "HR_SALARY_HISTORY", "edit/Add")
Resume Next

End Function

Private Sub clpCode_Change(Index As Integer)
If Index = 5 Then
    txtComment = clpCode(5)
End If
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
If Index = 0 Then
    Call Set_SalState
    Call Set_MarketLine_List
End If
If Index = 5 Then
    txtComment = clpCode(5)
End If
'Ticket #16276 - fglbNiagPhrs - contains the Pay Period per year and it was not getting refreshed on user selection
If Index = 4 And glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls
    fglbNiagPhrs = clpCode(4).Text
    Call setPayPeriodSalary
End If
End Sub

Private Sub clpGrid_LostFocus()
If Len(clpPostCode) = 0 Then Exit Sub
Call getJOB(clpPostCode, clpGrid)
If Set_Position(clpPostCode, False) Then
End If
If glbMulti Then Call Get_OrgSalary
End Sub

Private Sub clpPostCode_LostFocus()
If Len(clpPostCode) = 0 Then Exit Sub
If Set_Position(clpPostCode, False) Then
    lblBANDCode = fglbBAND
    dlpPosStDate = fglbSDate
    clpGrid = fglbGrid
    txtWHRS = fglbWhrs
    txtPayrollID = fglbPayrollID
    
    'Ticket #24482 - Town of Caledon - Using the VGroup field to store the Job's Division to create uniqueness between
    'multiple same Position and Start Date positions linked to Salary.
    If glbCompSerial = "S/N - 2182W" Then
        clpDiv = fglbDiv
        'Call setDivList(fglbJob$, fglbSDate)
    End If
Else
    lblBANDCode = ""
    dlpPosStDate = ""
    clpGrid = ""
    txtWHRS = ""
    txtPayrollID = ""
    
    'Ticket #24482 - Town of Caledon - Using the VGroup field to store the Job's Division to create uniqueness between
    'multiple same Position and Start Date positions linked to Salary.
    If glbCompSerial = "S/N - 2182W" Then
        clpDiv = ""
        'Call setDivList(fglbJob$, fglbSDate)
    End If
End If
Call getJOB(clpPostCode, clpGrid)

If glbMulti Then Call Get_OrgSalary
End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err

dlpDate(0).DataChanged = False
dlpDate(1).DataChanged = False
'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
''' Sam add July 2002 * Remove Binding Control
'rsDATA.CancelUpdate

fglbNew = False
Call Display_Value

'Ticket #15268 - The Grid Category description is not refreshing correctly.
If glbMultiGrid Then
    Call getJOB(clpPostCode, clpGrid)
End If

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_SALARY_HISTORY", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdChPos_Click()
clpPostCode.Enabled = True
dlpPosStDate.Enabled = True
txtWHRS.Enabled = True
clpGrid.Enabled = True
clpPostCode.SetFocus
'Need to be able to edit the hours for costed attendance Bryan Ticket #11870
'If chkCurrent.Value = 0 Then txtWHRS.Enabled = True Else txtWHRS.Enabled = False
End Sub

Private Sub cmdChPos_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMESALARY" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, xID
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant
Dim UpdateAudit  As Boolean
Dim UptSalaryDate As Date, orgEDate As Date
Dim DCurSHDate

If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Then  'Town of Aurora and timmis
    If chkCurrent <> 0 Then
        MsgBox "Please uncheck the Current Salary flag before deleting the record"
        Exit Sub
    End If
End If
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

'Do not allow deleting the last salary record if the integration is ON
If glbVadim Then
    'Check if Integration for Salary is ON
    If Not isTransfer(Salary) Then GoTo Cont_SalDel 'Exit Sub
    
    'Ticket #15070
    orgEDate = dlpDate(0)

    'Check if last Salary record
    Dim rsSal As New ADODB.Recordset
    
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & glbLEE_ID
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsSal.RecordCount = 1 Then
        MsgBox "This Salary record cannot be delete because it is the last Salary record.", vbInformation
        Exit Sub
    End If
    rsSal.Close
End If

Cont_SalDel:

DtTm = Now
DCurSHDate = CurSHDate()
fglHredsem = dlpDate(1).Text  '11/2/97 by Laura

If Trim(fglHredsem) <> "" Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If

ResetFlagAudit

'For Vadim and City of Niagara Falls only - Delete the existing entries Ticket #15542
'Other Vadim client do not have HR_VADIM_SY_INTERFACE table
If glbVadim And glbCompSerial = "S/N - 2276W" And (CVDate(Format(dlpDate(0), "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy"))) Then
    'Call procedure to delete from SY_INTERFACE and SY_INTERFACE_BATCH tables
    Call Delete_Existing_Vadim("D")
End If

'Comment enhacement - Ticket #16115
'City of Niagara Falls - Ticket #15542
If glbVadim And glbCompSerial = "S/N - 2276W" Then
    'Delete the salary record from Vadim's HR_EMP_HIST table storing the history of Rate changes
    Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID.Text, dlpDate(0), Val(lblSalaryGrade), "", clpPostCode, "D")
End If

xID = Data1.Recordset("SH_ID")
gdbAdoIhr001.BeginTrans
rsDATA.Delete 'gdbAdoIhr001.Execute "DELETE FROM HR_SALARY_HISTORY WHERE SH_ID=" & xID
gdbAdoIhr001.CommitTrans

If Not glbOracle And Not glbSQL Then Pause (0.5)
Data1.Refresh

If glbGP Then Call Salary_Integration(glbLEE_ID, , True, fglbNew, xID) 'George Mar 7,2006 #9965

prompt = False
Call cmdRecal_Click
prompt = True

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    'Update employee's Salary records with the correct Primary Position checkbox
    Call UpdatePrimaryPositionSalary(glbLEE_ID)
End If

Data1.Refresh

If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    Call Set_Current_Flag
End If

Call Display_Value

'Comment enhacement - Ticket #16115
'City of Niagara Falls - Ticket #15542
'Clear End Date in previous salary record in Vadim's HR_EMP_HIST table which is now current
If glbVadim And glbCompSerial = "S/N - 2276W" Then
    Call Update_VadimDB_HR_EMP_HISTORY(fglbPayrollID, dlpDate(0), "", "", "", "M", "")
End If

If OSalary <> medsalary And (chkCurrent Or Data1.Recordset.EOF) Then
    Call updBenefitForSalDEPN(glbLEE_ID) 'Jaddy 9/9/99
    If glbCompSerial = "S/N - 2380W" Then Call CalcPP 'VitalAire
    If glbCompSerial = "S/N - 2291W" Then Call updCompPlan(glbLEE_ID, Val(medsalary) - Val(OSalary), DCurSHDate)
End If

'Check if Non-factored salary has changed - update Salary Dependant Benefits
If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
    If ONFSalary <> medNFacSalary And (chkCurrent Or Data1.Recordset.EOF) Then
        Call updBenefitForSalDEPN(glbLEE_ID)
    End If
End If
If Not glbMediPay Then
    Call Employee_Master_Integration(glbLEE_ID)
End If

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

vbxTrueGrid.Refresh

If glbVadim Then
    Dim HRSalary As New Collection
    Dim HRChanges As New Collection
    
    UpdateAudit = False
    'Ticket #15070 - If it's a future dated record getting deleted, then this update to Vadim should be directed for future date as well.
    If CVDate(Format(orgEDate, "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy")) Then
        UptSalaryDate = orgEDate
    Else
        UptSalaryDate = Date
    End If

    If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
        If isChanged_Salary(HRSalary, OTOTAL, medTotal, True) Then UpdateAudit = True
    Else
        If isChanged_Salary(HRSalary, OSalary, medsalary, True) Then UpdateAudit = True
    End If
    If isChanged_Salary(HRSalary, OSalCD, lblSalCode) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbNiagPhrs, fglbDhrs, glbLEE_ID, txtPayrollID.Text)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, UptSalaryDate, fglbPhrs, fglbWhrs, glbLEE_ID, txtPayrollID.Text)
        End If
        'City of Kawartha Lakes - Pass Salary Grade to Probation Levels
        If glbCompSerial = "S/N - 2363W" Then
            If fglbNew Then
                If isChanged_Field(HRChanges, "", lblSalaryGrade, True) Then UpdateAudit = True
            Else
                If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then UpdateAudit = True
            End If
        Else
            If isChanged_Field(HRChanges, oGrade, lblSalaryGrade, True) Then Debug.Print "" ' do nothing for the audit transfer
        End If
    End If
    
    
    If isChanged_Field(HRChanges, OEDate, dlpDate(0)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, ONDate, dlpDate(1)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, SavPAYP, clpCode(4)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, OReason, clpCode(1)) Then UpdateAudit = True
    If isChanged_Field(HRChanges, oJob, clpPostCode) Then UpdateAudit = True
    If glbCompSerial = "S/N - 2373W" Then 'DMuskoka , ,  'Vailtech
        If isChanged_Field(HRChanges, OPremium, medPremium) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OTOTAL, medTotal) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OvGroup, txtVGroup) Then UpdateAudit = True
        If isChanged_Field(HRChanges, OVStep, txtVStep) Then UpdateAudit = True
    End If
    'Ticket #15070
    'Call Passing_Changes(HRChanges, Salary, "M", Date, glbLEE_ID, txtPayrollID.Text)
    Call Passing_Changes(HRChanges, Salary, "M", UptSalaryDate, glbLEE_ID, txtPayrollID.Text)
End If

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_SALARY_HISTORY", "Delete")
Call RollBack '28July99 js
End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
Dim SQLQ As String, X%
Dim Response%, Msg$, Title$, DgDef As Double

On Error GoTo Mod_Err

Call SET_UP_MODE

Actn = "M"
fglHredsem = dlpDate(1).Text
If Not Data1.Recordset.EOF Then
    If Not IsNull(Data1.Recordset("SH_JOB")) Then
        fglbJob$ = Data1.Recordset("SH_JOB")
    End If
End If
orgPosStDate = dlpPosStDate.Text

'orgSalary = Val(medSalary)
If glbFrench Then
    If IsNumeric(medsalary.Text) Then
        orgSalary1 = medsalary
        orgSalary = medsalary   'Release 8.0 - Ticket #22682
    End If
Else
    orgSalary1 = Val(medsalary)
    orgSalary = Val(medsalary)  'Release 8.0 - Ticket #22682
End If

'Hemu - essex
If glbFrench Then
    If IsNumeric(medAmtChng(1)) Then
        fglbAmtOld(1) = CCur(medAmtChng(1))
    End If
    If IsNumeric(medAmtChng(2)) Then
        fglbAmtOld(2) = CCur(medAmtChng(2))
    End If
    If IsNumeric(medAmtChng(3)) Then
        fglbAmtOld(3) = CCur(medAmtChng(3))
    End If
Else
    fglbAmtOld(1) = CCur(Val(medAmtChng(1)))
    fglbAmtOld(2) = CCur(Val(medAmtChng(2)))
    fglbAmtOld(3) = CCur(Val(medAmtChng(3)))
End If
'Hemu - essex

orgCurrent = chkCurrent
SavPAYP = clpCode(4).Text


SavSalcd = lblSalCode
    
'Release 8.0 - Logic Fix to calculate the % and Amount change with different Per
orgSalCD1 = lblSalCode


''If glbWFC And UnionExecNone Then
''    lblBANDCode = fglbBAND
''    optUserSys(0).Value = False: optUserSys(1).Value = True
''    optUserSys(0).Enabled = False: optUserSys(1).Enabled = True
''    mskCampa.Visible = optUserSys(1) And optUserSys(1).Visible
''    If Val(lblsalstate(1)) > 0 And Val(mskCampa) = 0 Then
''      If Val(lblCompaNum) > 0 And Val(lblCompaNum) < 999.99 Then
''        mskCampa = (Val(medSalary) / Val(lblCompaNum)) * 100
''      Else
''        mskCampa = Val(lblsalstate(1))
''      End If
''      mskCampa = Round2DEC(mskCampa)
''    End If
''End If

'clpCode(1).SetFocus
'If glbSetSal Or glbMulti Then clpPostCode.SetFocus

'clpCode(1).Enabled = True
'clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_SALARY_HISTORY", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String, Msg$
Dim X%
Dim orgMarketLine, orgSalCD
Dim xPayPeriod
Dim orgNextReviewDt

On Error GoTo AddN_Err

fglbNew = True

'Hemu - essex
fglbAmtOld(1) = 0
fglbAmtOld(2) = 0
fglbAmtOld(3) = 0
'Hemu - essex

'Ticket #21511
If glbSetSal Then
    Call CR_JobHis_Snap(False)
Else
    Call CR_JobHis_Snap(True)
End If

If Not Set_Position("", True) Then
    Msg$ = "No current position found "
    Msg$ = Msg$ & Chr(10) & "Please review position prior to updating salary."
    MsgBox Msg$
    Exit Sub
End If
If Not getJOB(fglbJob$, fglbGrid) Then   '- populates job items/grades
    If glbMultiGrid Then
        Msg$ = "Can not find Salary Details for current position and grid category."
        Msg$ = Msg$ & Chr(10) & "Please review position Master list and the Salary Details."
    Else
        Msg$ = "Can not find description for current position."
        Msg$ = Msg$ & Chr(10) & "Please review position Master list."
    End If
    MsgBox Msg$
    Exit Sub
Else    'Ticket #15268 - The Grid Category description is not refreshing correctly.
    If glbMultiGrid Then
        'Ticket #15708 - The Grid step is not getting populated when a new Salary rec is being entered for new Position
        'Call getJOB(clpPostCode, clpGrid)
        Call getJOB(fglbJob$, fglbGrid)
    End If
End If
If glbMulti And Not Data1.Recordset.EOF Then
    MsgBox "If necessary, edit the previous salary record to remove the current flag."
End If
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    Data1.Recordset.MoveFirst
    orgMarketLine = txtMarketLine
    
    If glbFrench Then
        If IsNumeric(medsalary.Text) Then
            orgSalary = medsalary
            orgSalary1 = medsalary
        End If
    Else
        orgSalary = Val(medsalary)
        orgSalary1 = Val(medsalary)
    End If
    orgSalCD = lblSalCode

    '''If glbMulti Then Call Get_OrgSalary 'Ticket #14354 commented by Frank, since this caused a problem to calculate the Percentage Change
Else
    orgMarketLine = ""
    orgSalary = 0
    orgSalary1 = 0
    orgSalCD = JobSnaps_Salary_Code$
End If
DoEvents
xPayPeriod = clpCode(4)

'Ticket #15723
'If glbLambton Then
'    'Ticket #15699
'    orgNextReviewDt = dlpDate(1).Text
'End If

fglbEmptyNew = (Data1.Recordset.BOF And Data1.Recordset.EOF)
Call Set_Control("B", Me)

'rsDATA.AddNew

If fglbReason$ = "NEWH" And fglbEmptyNew Then clpCode(1).Text = "NEWH"

Actn = "A"


lblCNum.Caption = "001"
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblWhrs = fglbWhrs#
txtWHRS = fglbWhrs#

clpPostCode.Text = fglbJob$
dlpPosStDate.Text = CVDate(fglbSDate)
clpGrid.Text = fglbGrid
txtPayrollID = fglbPayrollID
lblBANDCode = fglbBAND

'Ticket #24482 - Town of Caledon - Using the VGroup field to store the Job's Division to create uniqueness between
'multiple same Position and Start Date positions linked to Salary.
If glbCompSerial = "S/N - 2182W" Then
    clpDiv = fglbDiv
    Call setDivList(fglbJob$, fglbSDate)
End If

Call setGridList(fglbJob$)

'Ticket #14354 Added by Frank, since this caused a problem to calculate the Percentage Change
If glbMulti Then Call Get_OrgSalary

'If glbLinNewPosSal And glbLinamar Then 'Jaddy changed by linda asking, 8/20/01
If glbLinamar Then
    Call Set_NextReview
    If glbLinNewPosSal Then
        clpCode(1).Text = fglbReason$  'glbLinReasonCode
    End If
End If

If glbLambton Then
    If Len(xPayPeriod) > 0 Then
        clpCode(4) = xPayPeriod
    Else
        clpCode(4) = "26"
    End If
    'Ticket #15723
'    'Ticket #15699
'    dlpDate(1).Text = orgNextReviewDt
End If
'If glbMediPay Then
'    If Len(xPayPeriod) > 0 Then
'        clpCode(4) = xPayPeriod
'    End If
'End If
If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    Call Set_CommentFromUnion
End If

If glbCompSerial = "S/N - 2229W" Or glbCompSerial = "S/N - 2347W" Then 'Inscap Solution - Ticket # 8932 'Ticket #24504 SPC
    If Len(xPayPeriod) > 0 Then
        clpCode(4) = xPayPeriod
    End If
End If
If glbCompSerial = "S/N - 2397W" Then 'Ticket #15255
    clpCode(4) = "3880"
End If
lblSalaryGrade = "00"
If glbWFC Then
    clpCode(4) = getLastPayP
    clpCode(6) = getWFCCurrencyIndi(GetEmpData(glbLEE_ID, "ED_SECTION", "")) 'Ticket #29069 Franks 08/22/2016
End If
If glbCompSerial = "S/N - 2410W" Then 'Frontenac Ticket #19071
    clpCode(4) = "26"
End If

'WDGPHU - Ticket #17324
If glbCompSerial = "S/N - 2411W" Then
    txtPosGroup.Text = GetJobData(clpPostCode.Text, "JB_GRPCD", "")
End If

'Ticket #20652 - Town of Aurora
If glbCompSerial = "S/N - 2378W" Then
    lblPosGrp.Visible = True
    lblPosGrp.Caption = GetTABLDesc("JBGC", GetJobData(clpPostCode.Text, "JB_GRPCD", ""))
Else
    lblPosGrp.Caption = ""
    lblPosGrp.Visible = False
End If

'Ticket #24543 - Macaulay Child Development Centre
If glbCompSerial = "S/N - 2420W" Then
    clpCode(4) = "3BG2"
End If

'Ticket #19933 Franks 03/18/2011
If glbCompSerial = "S/N - 2382W" Then  'Samuel
    Call SetDefaultsSamuel
Else
    lblSalCode = orgSalCD
    
    'Release 8.0 - Logic Fix to calculate the % and Amount change with different Per
    orgSalCD1 = orgSalCD
End If
chkCurrent = glbMulti
medsalary = 0
SavPAYP = ""
SavSalcd = ""

If glbWFC Then 'Ticket #24184 Franks 09/11/2013
    If NewHireForms.count > 0 Then
        Call WFCHRSoftDispValues
        'Ticket #24695 Franks 11/28/2013 - begin
        clpCode(1).Text = "NEW"
        If lblSalCode.Caption = "A" Then clpCode(4).Text = "SM"
        If lblSalCode.Caption = "H" Then clpCode(4).Text = "W"
        'Ticket #24695 Franks 11/28/2013 - end
    Else
        'Ticket #25927 Franks 08/25/2014 - check if the HRSoft Salary Upt Flag is YES
        If IsFirstEmpSalary(glbLEE_ID) Then
            If WFCHRSoftMissNewhire(glbLEE_ID, "SF_UPT_SALARY") Then
                Call WFCHRSoftDispValues
                'Ticket #24695 Franks 11/28/2013 - begin
                clpCode(1).Text = "NEW"
                If lblSalCode.Caption = "A" Then clpCode(4).Text = "SM"
                If lblSalCode.Caption = "H" Then clpCode(4).Text = "W"
                'Ticket #24695 Franks 11/28/2013 - end
            End If
        End If
    End If
End If

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
clpCode(1).Enabled = True
clpCode(1).SetFocus

If clpCode(1).Text = "NEWH" Then

    fraSalary.Enabled = True
    For X% = 1 To 3
        medPercentChng(X%) = 0
        medPercentChng(X%).Enabled = False
        medAmtChng(X%) = 0
        medAmtChng(X%).Enabled = False
        If X% > 1 Then
            clpCode(X%).Enabled = False
        End If
    Next X%
Else
    medPercentChng(1).Enabled = True
    medAmtChng(1).Enabled = True
End If
comSalScale.ListIndex = 0
'If glbMediPay Then
'    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
'        clpCode(4).Enabled = False
'    End If
'End If
If glbWFC Then
    For X% = 0 To cmbMarketLine.ListCount
        If cmbMarketLine.List(X%) = orgMarketLine Then txtMarketLine = orgMarketLine
    Next
    'Ticket# 6962 Begin
    If clpCode(0).Visible Then
        clpCode(0) = glbEmpPlant
    End If
    If dlpDate(2).Visible Then
        dlpDate(2) = Format(Now, "SHORT DATE")
    End If
    'Ticket# 6962 Begin
End If
'If glbSetSal Or glbMulti Then clpPostCode.SetFocus

DoWFCGrids (True)

'Ticket #22045 - Call this function again because the above function (DoWFCGrids) is reseting the
'comSalScale list and so this function will repopulate again. I am excluding WFC because I
'believe they have special logic in the above function (DoWFCGrids).
If Not glbWFC Then
    Call getJOB(clpPostCode.Text, clpGrid.Text)
End If
''added by Bryan 24/Oct/05 Ticket#9607
'If glbCompSerial = "S/N - 2378W" Then
'    txtPayrollID = glbLEE_ID
'End If

Exit Sub

AddN_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SALARY_HISTORY", "Add")
Resume Next

End Sub

Private Sub Set_CommentFromUnion()
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & glbLEE_ID & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("JH_ORG")) Then
            clpCode(5).Text = rsTemp("JH_ORG")
        End If
    End If
    rsTemp.Close
    
End Sub
'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Function Set_SalaryGrade(xSalary As Double)
Dim SQLQ As String, X%
Dim xsSalary As Double
Dim strSalcode As String

If glbLambton Then 'Ticket# 6693
    If glbSetSal Then
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2259W" Then 'Oxford Ticket #17139
    'multi steps may have same rate, user wants to remember the step which they selected
    Exit Function
End If

If Len(fglbJob$) > 0 Then
    lblSalaryGrade = "00"
    xSalary = Round2DEC(xSalary)
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        If JobSnaps_Salary_Code$ = "H" Then
            If lblSalCode = "H" Then
                xsSalary = xSalary
            ElseIf lblSalCode = "M" Then
                If Val(lblWhrs) = 0 Then
                    xsSalary = 0
                Else
                    xsSalary = ((xSalary * 12) / Val(lblWhrs)) / 52
                End If
            ElseIf lblSalCode = "A" Then
                If Val(lblWhrs) = 0 Then
                    xsSalary = 0
                Else
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        xsSalary = (xSalary)
                    Else
                        xsSalary = (xSalary / Val(lblWhrs)) / 52
                    End If
                End If
            'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
            ElseIf lblSalCode = "D" Then
                If Val(lblWhrs) = 0 Then
                        xsSalary = 0
                    Else
                        If GetLeapYear(Year(Date)) Then
                            xsSalary = ((xSalary * 366) / Val(lblWhrs)) / 52
                        Else
                            xsSalary = ((xSalary * 365) / Val(lblWhrs)) / 52
                        End If
                        
                        'Ticket #17654 - formula correction
                        xsSalary = (xSalary / fglbDhrs)
                    End If
                End If
        ElseIf JobSnaps_Salary_Code$ = "A" Then
            If lblSalCode = "H" Then
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xsSalary = (xSalary)
                Else
                    xsSalary = (xSalary * Val(lblWhrs)) * 52
                End If
            ElseIf lblSalCode = "M" Then
                xsSalary = xSalary * 12
            ElseIf lblSalCode = "A" Then
                xsSalary = xSalary
            'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
            ElseIf lblSalCode = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xsSalary = (xSalary * 366)
                Else
                    xsSalary = (xSalary * 365)
                End If
                
                'Ticket #17654 - formula correction
                xsSalary = (xSalary / fglbDhrs) * Val(lblWhrs) * 52
            End If
        End If
        xsSalary = Round2DEC(xsSalary)
        If JobSnaps_PayScale(X%) <> 0 And xsSalary >= JobSnaps_PayScale(X%) Then
            lblSalaryGrade = Format(X%, "00")
        End If
    Next X%
End If
End Function

Private Function Set_SalaryGrade_French(xSalary)
Dim SQLQ As String, X%
Dim xsSalary
Dim strSalcode As String

If glbLambton Then 'Ticket# 6693
    If glbSetSal Then
        Exit Function
    End If
End If

If glbCompSerial = "S/N - 2259W" Then 'Oxford Ticket #17139
    'multi steps may have same rate, user wants to remember the step which they selected
    Exit Function
End If

If Len(fglbJob$) > 0 Then
    lblSalaryGrade = "00"
    xSalary = Round2DEC(xSalary)
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        If JobSnaps_Salary_Code$ = "H" Then
            If lblSalCode = "H" Then
                xsSalary = xSalary
            ElseIf lblSalCode = "M" Then
                If Val(lblWhrs) = 0 Then
                    xsSalary = 0
                Else
                    xsSalary = ((xSalary * 12) / Val(lblWhrs)) / 52
                End If
            ElseIf lblSalCode = "A" Then
                If Val(lblWhrs) = 0 Then
                    xsSalary = 0
                Else
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        xsSalary = (xSalary)
                    Else
                        xsSalary = (xSalary / Val(lblWhrs)) / 52
                    End If
                End If
            'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
            ElseIf lblSalCode = "D" Then
                If Val(lblWhrs) = 0 Then
                        xsSalary = 0
                    Else
                        If GetLeapYear(Year(Date)) Then
                            xsSalary = ((xSalary * 366) / Val(lblWhrs)) / 52
                        Else
                            xsSalary = ((xSalary * 365) / Val(lblWhrs)) / 52
                        End If
                        
                        'Ticket #17654 - formula correction
                        xsSalary = (xSalary / fglbDhrs)
                    End If
                End If
        ElseIf JobSnaps_Salary_Code$ = "A" Then
            If lblSalCode = "H" Then
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xsSalary = (xSalary)
                Else
                    xsSalary = (xSalary * Val(lblWhrs)) * 52
                End If
            ElseIf lblSalCode = "M" Then
                xsSalary = xSalary * 12
            ElseIf lblSalCode = "A" Then
                xsSalary = xSalary
            'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
            ElseIf lblSalCode = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xsSalary = (xSalary * 366)
                Else
                    xsSalary = (xSalary * 365)
                End If
                
                'Ticket #17654 - formula correction
                xsSalary = (xSalary / fglbDhrs) * Val(lblWhrs) * 52
            End If
        End If
        xsSalary = Round2DEC(xsSalary)
        If JobSnaps_PayScale(X%) <> 0 And xsSalary >= JobSnaps_PayScale(X%) Then
            lblSalaryGrade = Format(X%, "00")
        End If
    Next X%
End If
End Function

Sub cmdOK_Click()
Dim rsSal As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsBenCode As New ADODB.Recordset
Dim X, xID, xUpdCurrent
Dim vList As String
Dim SHMark
Dim xRevDate
Dim SQLQ As String
Dim a As Integer, Msg As String

On Error GoTo Add_Err
'Mostafa - Farmers Mutual
Dim xDiv
xDiv = GetDivisionCode

If glbCompSerial = "S/N - 2407W" Then
    If xDiv = "ND" Then
        clpCode(4).Text = "4024"
    End If
End If
'Hemu - Ticket #13830 - Remove the Password prompt from Salary screen.
''City of Timmins - Ticket #13207
'If glbCompSerial = "S/N - 2375W" And fglbNew <> True Then
'    'Check if Current checkbox is unchecked then do not prompt for Password
'    If orgCurrent <> chkCurrent And chkCurrent = False Then
'         'Save the changes and do not prompt for Password
'    Else
'        'Ask for the password
'        glbAccessPswd = False
'        frmAccessPswd.Show 1
'        If glbAccessPswd = False Then   'Access Denied
'            Call cmdCancel_Click
'            Exit Sub
'        End If
'    End If
'End If
'Hemu - End - Ticket #13830

If glbWFC And UnionExecNone Then
    lblBANDCode = fglbBAND
    'optUserSys(0).Value = False: optUserSys(1).Value = True
    'optUserSys(0).Enabled = False: optUserSys(1).Enabled = True
    'mskCampa.Visible = optUserSys(1) And optUserSys(1).Visible
    'If Val(lblsalstate(1)) > 0 And Val(mskCampa) = 0 Then
    '  If Val(lblCompaNum) > 0 And Val(lblCompaNum) < 999.99 Then
    '    mskCampa = (Val(medSalary) / Val(lblCompaNum)) * 100
    '  Else
    '    mskCampa = Val(lblsalstate(1))
    '  End If
    '  mskCampa = Round2DEC(mskCampa)
    'End If
End If

'Hemu - it was not saving the new the Group and Step if new items were added to the list
'commented it here and added these line below - after it assigns the value to txt fields
'vList = VGroupList
'vList = VStepList

If glbCompSerial = "S/N - 2436W" Then  'Family Day - Ticket #21152 Franks 04/01/2013
    If Not glbtermopen Then
        Call FamilayDaySalaryChange(glbLEE_ID, clpCode(4).Text, dlpDate(0).Text, fglbNew)
    End If
End If

If Not chkSalHist() Then Exit Sub

If clpPostCode.Enabled = True Then      'Laura nov 21, 1997
    If Not Chkpos() Then Exit Sub
End If

If glbWFC Then 'Ticket #19266
    Call AUDIT_NGS_TRANS
End If

If gsEMAIL_ONSALARY Then
    MailBody = ""
    If NewHireForms.count = 0 Then 'Non new hire
        If OSalary <> medsalary And OSalary > 0 And (fglbNew Or chkCurrent) And OSalary > 0 Then  'Only Salary Change

            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                MailBody = GetEmailBodyForSamuel(glbLEE_ID)
                MailBody = MailBody & "salary has been changed." & vbCrLf & vbCrLf
                'MailBody = MailBody & "New Salary: " & (Format(medsalary, "$#.00")) & "/" & comPayPer.Text & vbCrLf
                'MailBody = MailBody & "Reason: " & GetTablDesc("SDRC", clpCode(1)) & vbCrLf
                MailBody = MailBody & "Effective Date: " & dlpDate(0) & vbCrLf
            Else
                MailBody = "The Salary has been changed." & vbCrLf & vbCrLf
                MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                MailBody = MailBody & "New Salary: " & (Format(medsalary, "$#.00")) & "/" & comPayPer.Text & vbCrLf
                MailBody = MailBody & "Reason: " & GetTABLDesc("SDRC", clpCode(1)) & vbCrLf
                MailBody = MailBody & "Effective Date: " & dlpDate(0) & vbCrLf
                
                If glbCompSerial = "S/N - 2417W" Then  'Ticket #22334 - County of Perth
                    MailBody = MailBody & "Comments: " & txtComment.Text & vbCrLf
                End If
            End If
            'Screen.MousePointer = DEFAULT
            'Call imgEmail_Click
        End If
    End If
End If

'If Not chkSalHist() Then Exit Sub
Screen.MousePointer = HOURGLASS

glbChgTermReason = ""
If glbCompSerial = "S/N - 2351W" Or glbCompSerial = "S/N - 2387W" Then      'Burlington Tech
    'Bird Packaging Limited Ticket #13342
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbChgNewEmpnbr = lblEEID
    Screen.MousePointer = DEFAULT
    If SavPAYP <> clpCode(4).Text Then
        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
            frmMsgTerm.txtEmpNum.Enabled = False
            frmMsgTerm.Show 1
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

If (glbWFC And glbPlantCode = "GREN") Then   'Greensboro Ticket #15340
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbEESection = ""
    glbChgNewEmpnbr = lblEEID
    Screen.MousePointer = DEFAULT
    If SavPAYP <> clpCode(4).Text Then
        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
            Dim xComCodeChanged As Boolean
            xComCodeChanged = False
            If SavPAYP = "W" And (clpCode(4).Text = "SM" Or clpCode(4).Text = "M") Then
                xComCodeChanged = True
            End If
            If clpCode(4).Text = "W" And (SavPAYP = "SM" Or SavPAYP = "M") Then
                xComCodeChanged = True
            End If
            If xComCodeChanged Then
                frmMsgTerm.txtEmpNum.Enabled = False
                frmMsgTerm.Show 1
                If Len(glbChgTermDate) > 0 Then
                    glbChgTermReason = "TERM"
                End If
            End If
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

If glbCompSerial = "S/N - 2369W" Then    'TS Tech Ticket #11544
    'If Pay Period was changed from TEMP to other codes, that is a new hire in Payroll
    'Create a New hire data in HRAudit
    glbChgTermDate = ""
    glbChgTermReason = ""
    glbChgNewEmpnbr = lblEEID
    Screen.MousePointer = DEFAULT
    If SavPAYP <> clpCode(4).Text Then
        If SavPAYP = "TEMP" Then
            glbChgTermReason = "OTHER"
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

If glbCompSerial = "S/N - 2242W" Then    'London CCAC
    glbChgTermDate = ""
    glbChgPT = ""
    glbChgUseProfile = ""
    Screen.MousePointer = DEFAULT
    If SavPAYP <> clpCode(4).Text Then
        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
        Select Case SavPAYP
        Case "132"
            frmMsgConfirm.clpCode(0).Text = "CAS"
            frmMsgConfirm.clpCode(1).Text = "NO"
        Case "133"
            frmMsgConfirm.clpCode(0).Text = "FT"
            frmMsgConfirm.clpCode(1).Text = "YES"
        End Select
        frmMsgConfirm.Show 1
        If glbChgPT = "" Or glbChgUseProfile = "" Then
            Call cmdCancel_Click
            Exit Sub
        End If
        End If
    End If
    Screen.MousePointer = HOURGLASS
End If

'If this function is processing, it's disabled. ticket #10398
If glbDisabled Then GoTo end_line
glbDisabled = True

rsDATA.Requery
If fglbNew Then rsDATA.AddNew

Call UpdUStats(Me) ' update user's stats (who did it and when)

If (glbCompSerial = "S/N - 2290W") Or (glbCompSerial = "S/N - 2171W") Then
    Updstats(0).Text = Format(Now, "SHORT DATE")
Else
    Updstats(0).Text = Format(dlpDate(0).Text, "SHORT DATE")
End If

If Not glbWFC Then
    dlpDate(2).Text = Format(Now, "SHORT DATE")
End If
If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #22952 Franks 12/10/2012
    'do not use the following calculation
Else
    'added by Bryan 22/Sep/05 Ticket#9343
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
End If
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    'Ticket #27106 - They want to compute the Step # using the Salary Amount and not Total
    'Call Set_SalaryGrade(Val(medTotal))
    Call Set_SalaryGrade(Val(medsalary))
    txtVGroup = cboVGRoup
    txtVStep = cboVStep
Else
    'Not DNSSAB - just believe in what they have already entered
    'Not City of Timmins - Ticket #14699
    'Ticket #21821 - Only do this for Current Salary records. We should not be changing the history.
    If chkCurrent Then
        If glbCompSerial <> "S/N - 2388W" And glbCompSerial <> "S/N - 2375W" Then
            If glbFrench Then
                Call Set_SalaryGrade_French(medsalary)
            Else
                Call Set_SalaryGrade(Val(medsalary))
            End If
        End If
    End If
    
    'City of Timmins
    If glbCompSerial = "S/N - 2375W" Then
        'New Hires don't have a value in comsalscale. Ticket #10436
        Dim strScale As String
        If comSalScale.Text = "" Then
            strScale = 0
        Else
            strScale = comSalScale.Text
        End If
        If JobSnaps_PayScale(CInt(strScale)) <> Val(medsalary) Then
            MsgBox "Salary does not match the grid Step.", vbExclamation, "info:HR"
        End If
    End If
    
End If

'Ticket #25117 - Clients who do not have Write permission on info:HR folder will get this error 75: Path/File
'access error. I am disabling for all except the clients I noticed in this code are using these fields. These are
'North Perth, KN&V and DMuskoka
If glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2429W" Or glbCompSerial = "S/N - 2437W" Then
    vList = VGroupList
    vList = VStepList
End If

Call Set_COMPA
Call Set_WFC_COMPA

If Actn = "A" Or orgCurrent <> chkCurrent Then
    xUpdCurrent = True
End If

'Comment by Frank Ticket #13851
'If glbCompSerial = "S/N - 2214W" Then
'    Dim xToDate
'    If IsDate(dlpDate(1).Text) Then
'        xToDate = dlpDate(1)
'    Else
'        xToDate = DateAdd("D", -1, DateAdd("YYYY", 1, CVDate(dlpDate(0).Text)))
'    End If
'    If Actn = "A" Then
'        Call ChangeOtherEarnAmount(lblEEID, medsalary, "A", dlpDate(0).Text, xToDate)
'    End If
'    If Actn = "M" And chkCurrent Then
'        Call ChangeOtherEarnAmount(lblEEID, medsalary, "M", dlpDate(0).Text, xToDate)
'    End If
'End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    'Current Salary is turned ON, set the corresponding old Salary's Current flag OFF
    If orgCurrent <> chkCurrent Or chkCurrent Then
        Call SetCurrentSalary_OFF(glbLEE_ID, clpPostCode.Text, dlpPosStDate.Text)
    End If
End If

Call Set_Control("U", Me, rsDATA)

If Val(lblSalaryGrade) = 0 Then rsDATA!SH_GRADE = "00"
rsDATA!sh_compa_user = IIf(optUserSys(0), "", "U")

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    'gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    'gdbAdoIhr001X.CommitTrans
    rsDATA.Resync
    xID = rsDATA("SH_ID")
Else
    'gdbAdoIhr001.BeginTrans
    rsDATA.Update
    'gdbAdoIhr001.CommitTrans
    rsDATA.Requery
    xID = rsDATA("SH_ID")
End If

If xUpdCurrent Then
    Call Set_Current_Flag
End If

xRevDate = fglHredsem

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    'Update employee's Salary records with the correct Primary Position checkbox
    Call UpdatePrimaryPositionSalary(glbLEE_ID)
End If

'Ticket #28595 - Prompt if to Update employee's Attendance records as welll
If NewHireForms.count = 0 Then 'Non new hire
    If OSalary <> medsalary And OSalary > 0 And (fglbNew Or chkCurrent) And OSalary > 0 Then  'Only Salary Change
        'Msg = "Do you want to update employee's Attendance records with New Salary as well?"
        Msg = "Do you want to update the new Salary information on the employee's Attendance records from " & Format(dlpDate(0), "mmm dd, yyyy") & " forward? "
        a% = MsgBox(Msg, 36, "Update Salary on Attendance Records")
        If a% = 6 Then
            Call Update_Attendance_SalaryInfo1(glbLEE_ID, clpPostCode.Text, dlpDate(0), medsalary, lblSalCode)
        End If
    End If
End If

Data1.Refresh
DoEvents
prompt = False
'Call cmdRecal_Click 'Ticket #14354 commented by Frank, since this caused a problem to calculate the Percentage Change

DoEvents
prompt = True

Data1.Recordset.Find "SH_ID=" & xID
Data1.Refresh

fglHredsem = xRevDate

'If glbMediPay Then    'MediPay
'    If SavPAYP <> clpCode(4).Text Then
'        If Len(SavPAYP) > 0 And Len(clpCode(4).Text) > 0 Then
'            If glbCompSerial = "S/N - 2242W" Then
'                Call UpdatePTAdministeredBy(glbChgPT, glbChgUseProfile)
'            End If
'            Call Employee_Transfered_MediPay(glbLEE_ID & "|" & SavPAYP)  ' for #8189
'        End If
'    End If
'End If

glbFlag_BenefitForSalDEPN = False
If OSalary <> medsalary And chkCurrent Then
    Call updBenefitForSalDEPN(glbLEE_ID) 'Jaddy 9/9/99
    
    If glbCompSerial = "S/N - 2380W" Then Call CalcPP  'VitalAire Ticket #11737
    
    If glbCompSerial = "S/N - 2291W" Then Call updCompPlan(glbLEE_ID, Val(medsalary) - Val(OSalary), dlpDate(0).Text)
    
    If glbWFC Then 'Ticket #23247 Franks 04/23/2013
        Call WFC_Salary_US_Ben(glbLEE_ID)
    End If
End If

'Check if Non-factored salary has changed - update Salary Dependant Benefits
If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
    If ONFSalary <> medNFacSalary And (chkCurrent Or Data1.Recordset.EOF) Then
        Call updBenefitForSalDEPN(glbLEE_ID)
    End If
End If

'Ticket #22682: Release 8.0 - Set older Salary Review Follow Up records as Completed first if uncompleted
'follow up records are found for Salary, before adding a new follow up record.
If fglbNew And NewHireForms.count = 0 Then
    glbFollowUpList = "SREV"
    If Older_FollowUp_Records_Found(glbFollowUpList) Then
        frmFollowUpList.Show 1
    End If
End If

If chkCurrent Then
    If Not updFollow("U") Then GoTo end_line 'Exit Sub
End If

'moved to after updFollow by Bryan Ticket#9294
Call Display_Value

DoEvents

'medipay doesn't need the employee master tansfer here
Dim saveMedipay

'Ticket #20410 Franks 06/03/2011, salary change not go to Medipay,
'reason: "glbMediPay = False", so comment it out
''saveMedipay = glbMediPay: glbMediPay = False
If Not glbMediPay Then
    Call Employee_Master_Integration(glbLEE_ID)
End If

If glbGP Then 'George Mar 7,2006 #9965
    Call Salary_Integration(glbLEE_ID, , False, fglbNew, xID) 'George Mar 7,2006 #9965
End If

If glbMediPay Then 'Ticket #14752
    Call Salary_Integration(glbLEE_ID)
End If

''medipay doesn't need the employee master tansfer here
'Dim saveMedipay
'saveMedipay = glbMediPay: glbMediPay = False
'If Not glbMediPay Then
'    Call Employee_Master_Integration(glbLEE_ID)
'End If

'Ticket #20410 Franks 06/03/2011
''glbMediPay = saveMedipay

'Ticket #18790 - Update EEO record
If glbEmpCountry = "U.S.A." Then
    If fglbNew Then
        Call uptEEO_Fields(glbLEE_ID, "Update")
    Else
        If Not oJob = clpPostCode.Text Then
            Call uptEEO_Fields(glbLEE_ID, "Update")
        End If
    End If
End If

fglbEmptyNew = False

'Ticket #25152: Macaulay Child Development Centre - Move to Performance Review screen if New Position/New Salary
'Trying to save the value for the call later in the end to call Performance Review screen.
If glbCompSerial <> "S/N - 2420W" Then
    fglbNew = False
End If

glbDisabled = False

Call SET_UP_MODE

Screen.MousePointer = DEFAULT

If glbOttawaCCAC Then
    If chkCurrent Then
        If clpCode(4).Text = "E" Then
            Dim oWHRS, oPHRS
            oWHRS = GetJHData(glbLEE_ID, "JH_WHRS", 0)
            oPHRS = GetJHData(glbLEE_ID, "JH_PHRS", 0)
            If oWHRS = 0 And oPHRS = 0 Then
                MsgBox "Please enter Hours/Week and Hours/Pay Period on Emplopee Position screen."
                Exit Sub
            Else
                If oWHRS = 0 Then
                    MsgBox "Please enter Hours/Week on Emplopee Position screen."
                    Exit Sub
                End If
                If oPHRS = 0 Then
                    MsgBox "Please enter Hours/Pay Period on Emplopee Position screen."
                    Exit Sub
                End If
            End If
        End If
    End If
End If
    
If glbCompSerial = "S/N - 2351W" Or glbCompSerial = "S/N - 2369W" Or glbCompSerial = "S/N - 2387W" Then      'Burlington Tech 'TS Tech 'Bird Packaging
    'rehiring is using most of the fields in HRAUDIT. Ticket#9899
    rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    If Len(glbChgTermReason) > 0 Then
        Call TermRehireAudit(rsTA)
    End If
    rsTA.Close
End If

If (glbWFC And glbPlantCode = "GREN") Then 'Ticket #15340
    If Len(glbChgTermReason) > 0 Then
        rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Call TermRehireAudit(rsTA)
        rsTA.Close
    End If
End If
        
If glbWFC Then 'Ticket #27774 Franks 12/30/2015
    If NewHireForms.count > 0 Then 'new hire only
        Call WFCAutoPerformance(glbLEE_ID)
    End If
End If

If gsEMAIL_ONSALARY Then
    If Len(MailBody) > 0 Then
        If glbFlag_BenefitForSalDEPN Then
            MailBody = MailBody & "The Salary dependent benefits has changed too." & vbCrLf
        End If
        Screen.MousePointer = DEFAULT
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352
            Call EmailSendingForSamuel
        Else
            Call imgEmail_Click
        
            'Release 8.1 - Email will be sent on Benefit changes as well.
            If gsEMAIL_ONBENEFIT Then
                If glbFlag_BenefitForSalDEPN And glbBenChanged <> "" Then
                    'Send Email
                    MailBody = "The Salary Dependent Benefit Update:" & vbCrLf & vbCrLf

                    MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
                    MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
                                                                    
                    'Following Benefits were Updated
                    If glbBenChanged <> "" Then
                        MailBody = MailBody & vbCrLf & "The following Benefit(s) got updated : " & vbCrLf
                        
                        'Retrieve the Benefits updated to get the Effective Date and Benefit Code Description
                        SQLQ = "SELECT BF_BCODE, BF_EDATE FROM HRBENFT WHERE BF_EMPNBR = " & glbLEE_ID
                        SQLQ = SQLQ & " AND BF_BCODE IN ('" & Replace(glbBenChanged, ",", "','") & "')"
                        rsBenCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
                        Do While Not rsBenCode.EOF
                            'Mail Body
                            MailBody = MailBody & vbTab & " - " & GetTABLDesc("BNCD", rsBenCode("BF_BCODE")) & " with Effective Date: " & Format(rsBenCode("BF_EDATE"), "SHORT DATE") & vbCrLf
                            rsBenCode.MoveNext
                        Loop
                        rsBenCode.Close
                        Set rsBenCode = Nothing
                    End If
                    Call imgEmailBenefit_Click
                    
                    Screen.MousePointer = DEFAULT
                End If
            End If
        
        End If
    End If
End If

'Ticket #25152: Macaulay Child Development Centre - Move to Performance Review screen if New Position/New Salary
If gSec_Inq_Performance And glbCompSerial = "S/N - 2420W" And NewHireForms.count = 0 And fglbNew Then
    frmEPERFORM.Show
End If
fglbNew = False

Call NextForm

If glbWFC Then 'Ticket #25927 Franks 08/26/2014 - for hrsoft missing position process
    'If NewHireForms.count = 0 Then
        If glbCandidate > 0 Then
             Call WFCHRSoftProcUpt("frmESALARY")
        End If
    'End If
End If

end_line:
Exit Sub

Add_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdPerform_Click()
'Unload frmEPERFORM
'glbSetPer = glbSetSal
'frmEPERFORM.Show
'Unload Me
'End Sub

Private Sub cmdPerform_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPosition_Click()
'Unload frmEPOSITION
'glbSetPos = glbSetSal
'frmEPOSITION.Show
'Unload Me

'End Sub

Private Sub cmdPosition_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

RHeading = lblEEName & "'s Salary History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 2
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "rgridSal.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HR_SALARY_HISTORY.SH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    xReport = glbIHRREPORTS & "rgridSa2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_SALARY_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Salary History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 2
            Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next
    End If
    xReport = glbIHRREPORTS & "rgridSal.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{HR_SALARY_HISTORY.SH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    xReport = glbIHRREPORTS & "rgridSa2.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.SelectionFormula = "{Term_SALARY_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
    
End If


Me.vbxCrystal.Action = 0

'cmdPrint.Enabled = True

End Sub



Private Sub CodeEnter(Indx As Integer)

If fglbReason$ <> "NEWH" And Indx < 4 Then
    If Len(clpCode(Indx).Text) > 0 Then
        medPercentChng(Indx).Enabled = True
        medAmtChng(Indx).Enabled = True
    Else
        medPercentChng(Indx) = 0
        medPercentChng(Indx).Enabled = False
        medAmtChng(Indx) = 0
        medAmtChng(Indx).Enabled = False
    End If
End If

End Sub

Private Sub cmdRecal_Click()
Dim xSalary
Dim Msg, a%

If prompt <> False Then
    Msg = "Are You Sure You Want To Recalculate the Percentage and Amount Change(s) For This Employee? "
    Msg = Msg & Chr(10) & Chr(10) & " This Action Will Ignore Records Have Multi-Reason. "
    a% = MsgBox(Msg, 36, "Confirm Recalulate")
    If a% <> 6 Then Exit Sub
End If

Data1.Refresh 'added by Bryan 05-08-05 Ticket #9063
With Data1.Recordset
    If .EOF Then Exit Sub
    xSalary = 0
    .MoveLast
    Do Until .BOF
        If IsNull(.Fields("SH_SREAS2")) And IsNull(.Fields("SH_SREAS3")) Then
            If xSalary = 0 Then
                .Fields("SH_SALPC1") = 1
                .Fields("SH_SALCHG1") = 0
            Else
                .Fields("SH_SALPC1") = (.Fields("SH_SALARY") - xSalary) / xSalary
                .Fields("SH_SALCHG1") = (.Fields("SH_SALARY") - xSalary)
            End If
            .Update
        End If
        xSalary = .Fields("SH_SALARY")
        .MovePrevious
    Loop
    .MoveFirst
End With

Call Set_COMPA
If prompt <> False Then
    DoEvents
    Data1.Refresh
    If Not glbSQL And Not glbOracle Then Call Pause(0.3)
    Data1.Refresh
    DoEvents
    If Not glbSQL And Not glbOracle Then Call Pause(0.3)
    Display_Value
    DoEvents
    Screen.MousePointer = DEFAULT
End If

End Sub

Private Sub cmdRecal_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdTranDate_Click()
        glbAccessPswd = False
        frmAccessPswd.Show 1
        If glbAccessPswd = False Then   'Access Denied
            Exit Sub
        End If
        dlpDate(2).Enabled = True
End Sub

Private Sub comPayPer_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comPayPer_LostFocus()
Dim z%

If comPayPer.ListIndex = 0 Then
    lblSalCode.Caption = "A"
ElseIf comPayPer.ListIndex = 1 Then
    lblSalCode.Caption = "H"
'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
'ElseIf glbCompSerial = "S/N - 2282W" And comPayPer.ListIndex = 3 Then
ElseIf comPayPer.ListIndex = 3 Then     'Ticket #17654
    lblSalCode.Caption = "D"
Else
    lblSalCode.Caption = "M"
End If

'Release 8.0 - Ticket #22682
Call medSalary_LostFocus
Call comSalScale_Click

If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
    z% = getJOB(clpPostCode.Text, clpGrid.Text)
End If
End Sub


Private Sub FIND_JOB()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
On Error GoTo Job_Err
Dim rsJOBs As New ADODB.Recordset

Screen.MousePointer = HOURGLASS
'SQLQ = "SELECT JB_CODE FROM HRJOB"
SQLQ = "SELECT TOP 10 JB_CODE FROM HRJOB" 'Ticket #27983 Franks 02/10/2016
rsJOBs.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If rsJOBs.EOF And rsJOBs.BOF Then
    Msg = "No Job descriptions found" & Chr(10)
    Msg = Msg & "You will require authority to add one to continue"
    MsgBox Msg
End If
'If Not IsNull(rsJOBs("JB_BAND")) Then
'    fglbBAND = IIf(IsNull(rsJOBs("JB_BAND")), "", rsJOBs("JB_BAND"))
'    lblBANDCode = fglbBAND
'End If
rsJOBs.Close

Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Jobs", "HRJOB", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next
 
End Sub

Private Sub CR_JobHis_Snap(xCurrent As Boolean)
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo JobHis_Err

Screen.MousePointer = HOURGLASS

'Added this line because Salary screen shows position as Unassigned for the existing Position when a new Position is
'added with New Salary option. This line will make sure the Position lookup is loaded with all the Positions the employee
'ever had. As it is the Salary is Position screen dependent so only Positions listed on Position screen should be in
'the lookup.
xCurrent = False

If glbtermopen Then
    SQLQ = "Select * from Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    'Ticket #21511
    If xCurrent Then
        SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    End If
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    Set dynaJobHIS = Nothing
    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
    'Ticket #21511
    If xCurrent Then
        SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    End If
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    Set dynaJobHIS = Nothing
    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If

xDefPosition = ""
If Not dynaJobHIS.EOF Then
    fglbJobList = ""
    dynaJobHIS.MoveFirst
    
    Do Until dynaJobHIS.EOF
        If Not IsNull(dynaJobHIS!JH_JOB) Then
            fglbJobList = fglbJobList & dynaJobHIS!JH_JOB & ","
        End If
        If Not IsNull(dynaJobHIS!JH_POSITION_CONTROL) Then 'North Perth Ticket #19209 Franks 05/18/2011
            If dynaJobHIS!JH_POSITION_CONTROL = "YES" Then
                xDefPosition = dynaJobHIS!JH_JOB
            End If
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
Private Sub Set_NextReview()
Dim EMP_Snap As New ADODB.Recordset
Dim SQLQ, xDATE, xLinDate, NewDate, dtY1%, dtY2%
    'Get Linamar Start Date
    SQLQ = "Select ED_EMPNBR,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    EMP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not (EMP_Snap.BOF And EMP_Snap.EOF) Then
        xLinDate = EMP_Snap("ED_DOH")
        If IsDate(xLinDate) Then
            xDATE = CurSHDate()
            
            If IsDate(xDATE) Then
                dtY1% = DateDiff("yyyy", CVDate(xLinDate), CVDate(xDATE))
                NewDate = DateAdd("yyyy", (dtY1% + 1), CVDate(xLinDate))
            Else
                NewDate = DateAdd("m", 3, CVDate(xLinDate))
            End If
            dlpDate(1) = NewDate
        End If
    End If
    EMP_Snap.Close
    
End Sub

Private Function CurSHDate()
Dim SQLQ As String
Dim HRSH_Snap As New ADODB.Recordset

CurSHDate = 0    ' returns 0 if no found records

On Error GoTo JS_Err

SQLQ = "Select * from HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND SH_CURRENT <>0"
'Town of Aurora or City of Timmins or City of Kawartha Lakes
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2363W" Then
    SQLQ = SQLQ & " ORDER BY SH_SALARY"
ElseIf glbMulti And glbVadim Then
    SQLQ = SQLQ & " AND SH_PAYROLL_ID='" & txtPayrollID.Text & "'"
    SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
ElseIf glbMulti Then
    SQLQ = SQLQ & " AND SH_JOB='" & clpPostCode.Text & "'"
    SQLQ = SQLQ & " ORDER BY SH_EDATE DESC"
End If
HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If HRSH_Snap.BOF And HRSH_Snap.EOF Then
    OSalary = 0
    OSalCD = ""
    oJob = ""
    OEDate = "01/01/01"
    ONDate = "01/01/01"
    OReason = ""
    OLambtonJob = ""
    SavPAYP = ""
    OldPAYP = ""
    oGrade = "00"
    OPremium = "": OTOTAL = "": OvGroup = "": OVStep = ""
    
    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
        ONFSalary = 0
    End If
Else
    'Not Town of Aurora and City of Timmins and not City of Kawartha Lakes
    If Not glbCompSerial = "S/N - 2378W" And Not glbCompSerial = "S/N - 2375W" And Not glbCompSerial = "S/N - 2363W" Then
        If fglbNew Then
            If glbMulti And glbVadim Then
                If HRSH_Snap("SH_PAYROLL_ID") = Data1.Recordset("SH_PAYROLL_ID") Then
                    HRSH_Snap("SH_CURRENT") = 0
                    HRSH_Snap.Update
                End If
            End If
        End If
    End If
    CurSHDate = HRSH_Snap("SH_EDATE")
    OSalary = HRSH_Snap("SH_SALARY")
    OSalCD = HRSH_Snap("SH_SALCD")
    oJob = HRSH_Snap("SH_JOB")
    OEDate = HRSH_Snap("SH_EDATE")
    ONDate = HRSH_Snap("SH_NEXTDAT")
    OReason = HRSH_Snap("SH_SREAS1")
    OLambtonJob = Left(HRSH_Snap("SH_GRID"), 1) & HRSH_Snap("SH_JOB") & Mid(HRSH_Snap("SH_GRID"), 2)
    SavPAYP = HRSH_Snap("SH_PAYP")
    OldPAYP = SavPAYP
    oGrade = HRSH_Snap("SH_GRADE")
    If glbCompSerial = "S/N - 2373W" Then 'Muskoka
        OPremium = HRSH_Snap("SH_PREMIUM"): OTOTAL = HRSH_Snap("SH_TOTAL")
        OvGroup = HRSH_Snap("SH_VGROUP"): OVStep = HRSH_Snap("SH_VSTEP")
    End If
    
    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
        ONFSalary = HRSH_Snap("SH_NFAC_SALARY")
    End If
End If

HRSH_Snap.Close
Exit Function

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SALARY History Snap", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Function

Function EERetrieve()
Dim SQLQ As String
Dim rs As New ADODB.Recordset

EERetrieve = False

On Error GoTo EERError
    If glbCompSerial = "S/N - 2259W" Then 'Added by Bryan 11/07/05 Ticket #8857 Oxford
        If glbtermopen Then
            SQLQ = "Select ED_SECTION FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        Else
            SQLQ = "Select ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        End If
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
        If rs("ED_SECTION") = "Y" Then
            glbMulti = True
            lblPayID.Visible = True
            txtPayrollID.Visible = True
        Else
            glbMulti = False
            lblPayID.Visible = False
            txtPayrollID.Visible = False
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
            lblPayID.Visible = True
            txtPayrollID.Visible = True
        Else
            glbMulti = False
            lblPayID.Visible = False
            txtPayrollID.Visible = False
        End If
        rs.Close
        Set rs = Nothing
        SQLQ = ""
    End If


If glbtermopen Then
    If glbCompSerial = "S/N - 2191W" Then 'A.E.F.O.
        'vbxTrueGrid.Columns(5).NumberFormat = "0.0"
        vbxTrueGrid.Columns(6).NumberFormat = "0.0"
    End If
    If glbOracle Then
        SQLQ = SQLQ & "SELECT Term_SALARY_HISTORY.*, SH_GRADE AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
    Else
        SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW, JB_GRPCD FROM Term_SALARY_HISTORY "
        'Ticket #20716 Franks 07/29/2011
        'SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
        SQLQ = SQLQ & " LEFT JOIN HRJOB ON Term_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
    End If
    
    SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
Else
    If glbCompSerial = "S/N - 2191W" Then 'A.E.F.O.
        SQLQ = SQLQ & " SELECT *,IIF(ISNULL(JB_DESCR2),SH_GRADE,IIF(JB_DESCR2<>'.5' OR SH_GRADE='00', VAL(SH_GRADE),(VAL(SH_GRADE)+1)/2)) AS SH_GRADESHOW "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
        SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
        SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
        'vbxTrueGrid.Columns(5).NumberFormat = "0.0"
        vbxTrueGrid.Columns(6).NumberFormat = "0.0"
    Else
        If glbOracle Then
             SQLQ = SQLQ & "SELECT HR_SALARY_HISTORY.*, SH_GRADE AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
        Else
             SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW, JB_GRPCD FROM HR_SALARY_HISTORY "
             SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
        End If
        SQLQ = SQLQ & "WHERE SH_EMPNBR = " & glbLEE_ID
        
    End If
End If
SQLQ = SQLQ & " ORDER BY "

'Ticket #21511 - County of Oxford - since they are able to switch between multi and non-multi, they are
'seeing an issue with sort order, so this will fix it.
If glbCompSerial = "S/N - 2259W" Then
    SQLQ = SQLQ & " SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_EDATE DESC"
ElseIf glbMulti Then
    SQLQ = SQLQ & " SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_EDATE DESC"
Else
    SQLQ = SQLQ & " SH_EDATE DESC, SH_ID DESC, SH_CURRENT " & IIf(glbSQL, "DESC", "")
End If

If glbCompSerial = "S/N - 2351W" Then   'Burlington Tech.
    'vbxTrueGrid.Columns(5).Visible = False
    vbxTrueGrid.Columns(6).Visible = False
End If

Data1.RecordSource = SQLQ
Data1.Refresh
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    Data1.Recordset.MoveFirst
    Data1.Recordset.Find "SH_CURRENT<>0"
End If
If glbWFC Then
    'Get Employee Plant code
    Call GetPlantCode
End If
EERetrieve = True

Call Display_Value

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Salary", "HR_SALARY_HISTORY", "SELECT")
Unload Me
Resume Next
Exit Function

End Function

Private Sub comSalScale_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comSalScale_Click()
Dim ssalary, HoursPerWeek!
Dim z%

If fglbGridType = 0.5 And Val(comSalScale) > 0 Then
    lblSalaryGrade = Format((Val(comSalScale) * 2 - 1), "00")
Else
    lblSalaryGrade = Format(Val(comSalScale), "00")
End If

If glbLambton Then 'Ticket# 6693
    If glbSetSal Then
        Exit Sub
    End If
End If

If lblSalaryGrade <> "00" And comSalScale.Enabled = True Then
    HoursPerWeek! = Val(lblWhrs)
    
    
    ''Ticket #26837 - Do not reset the Salary Grade/Step when the Salary is not Current or New
    'If fglbNew Or chkCurrent Then
        ssalary = JobSnaps_PayScale(Val(lblSalaryGrade))
    'Else
    '    ssalary = medsalary
    'End If
    
    If JobSnaps_Salary_Code$ = "H" Then
        If lblSalCode = "H" Then
            medsalary = Round2DEC(ssalary)
        'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
        ElseIf lblSalCode = "D" Then
            If IsDate(dlpDate(0)) Then
                If GetLeapYear(Year(dlpDate(0))) Then
                    medsalary = Round2DEC((ssalary * HoursPerWeek!) * 366) / 52
                Else
                    medsalary = Round2DEC((ssalary * HoursPerWeek!) * 365) / 52
                End If
                
                'Ticket #17654 - formula correction
                medsalary = Round2DEC((ssalary * fglbDhrs))
            End If
        ElseIf lblSalCode = "A" Then
            If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                medsalary = Round2DEC(ssalary)
            Else
                medsalary = Round2DEC((ssalary * HoursPerWeek!) * 52)
            End If
        ElseIf lblSalCode = "M" Then
            medsalary = Round2DEC(((ssalary * HoursPerWeek!) * 52) / 12)
        End If
    ElseIf JobSnaps_Salary_Code$ = "A" Then
        If lblSalCode = "H" Then
            If HoursPerWeek! = 0 Then
                medsalary = 0
            Else
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    medsalary = Round2DEC(ssalary)
                ElseIf glbCompSerial = "S/N - 2388W" Then   'DNSSAB Ticket #14475
                    If IsNumeric(JobSnaps_Salary_FTEHrs) And Val(JobSnaps_Salary_FTEHrs) > 1 Then
                        medsalary = Round2DEC(ssalary / JobSnaps_Salary_FTEHrs)
                    Else
                        MsgBox "There is no FTE Hours/Year for Position Code '" & clpPostCode & "' " & Chr(10) & "Please go to Position Master screen to enter this field. "
                        Exit Sub
                    End If
                Else
                    medsalary = Round2DEC((ssalary / HoursPerWeek!) / 52)
                End If
            End If
        'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
         ElseIf lblSalCode = "D" Then
            If IsDate(dlpDate(0)) Then
                If GetLeapYear(Year(dlpDate(0))) Then
                    medsalary = Round2DEC(ssalary * 366)
                Else
                    medsalary = Round2DEC(ssalary * 365)
                End If
                
                'Ticket #17654 - formula correction
                medsalary = Round2DEC(((ssalary / 52) / HoursPerWeek!) * fglbDhrs)
            End If
        ElseIf lblSalCode = "A" Then
            medsalary = Round2DEC(ssalary)
        ElseIf lblSalCode = "M" Then
            medsalary = Round2DEC(ssalary / 12)
        End If
    End If
    If glbFrench Then
        medsalary = Round2DEC(medsalary)
    Else
        medsalary = Round2DEC(Val(medsalary))
    End If
    Call setPercent
End If
End Sub

Private Sub dlpDate_Change(Index As Integer)
    If Index = 2 Then
        If IsDate(dlpDate(2)) Then
            dlpDate(2) = Format(dlpDate(2), "SHORT DATE")
        End If
    End If
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMESALARY"

'Ticket #24048 - Moving this to Form Load because, if adding a New Salary after adding a New Position, if the user
'does a lost focus on the Form and then comes back, the fglbNew is set to False causing the previous salary to be
'overwritten instead of new salary being added.
'fglbNew = False

flgloaded = True
glbDisabled = False
Call SET_UP_MODE
'Me.cmdModify_Click
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMESALARY"
End Sub

Sub Form_Load()
flagFrmLoad = True 'carmen may 00
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
On Error GoTo Err_Deal

'Ticket #24048 - Moving from Form Activate because, if adding a New Salary after adding a New Position, if the user
'does a lost focus on the Form and then comes back, the fglbNew is set to False causing the previous salary to be
'overwritten instead of new salary being added.
fglbNew = False

fraSalary2.BorderStyle = 0

If glbVadim Then
    lblPayID.FontBold = True
End If
If glbLambton Then
    lblLambtonJob.Visible = True
    txtLambtonJob.Visible = True
End If
If glbMulti Then
    If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/22/2014 Franks
        'dont show payroll id
    Else
        lblPayID.Visible = True
        txtPayrollID.Visible = True
    End If
End If
If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
End If

If glbWFC Then
    clpPostCode.TextBoxWidth = 1315 '1180 'Ticket #25911 Franks 11/10/2014
    dlpDate(2).DataField = "SH_TRANSDATE"
    txtFiscalYear.DataField = "SH_FISCALYEAR"
    clpCode(0).DataField = "SH_SECTION"
    txtMarketLine.DataField = "SH_MARKETLINE"
    'Ticket #29069 Franks 08/22/2016
    clpCode(6).DataField = "SH_CURRENCYINDI"
    clpCode(6).Visible = True
    lblCurrencyIndicator.Visible = True
Else
    dlpDate(2).Enabled = False
    clpPostCode.TextBoxWidth = 1315 'Ticket #26726 Franks 06/15/2015
End If
'added by Bryan 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    fraSalary.Height = 1515
    medPremium.DataField = "SH_PREMIUM"
    medTotal.DataField = "SH_TOTAL"
    txtVGroup.DataField = "SH_VGROUP"
    txtVStep.DataField = "SH_VSTEP"
    
    'Ticket #19113 - Hide Vailtech fields because they are not using Vailtech anymore
    lblTitle(19).Visible = False
    lblTitle(20).Visible = False
    cboVGRoup.Visible = False
    cboVStep.Visible = False
    
    'Ticket #24565 - They want this to be only mandatory when New Hire
    If NewHireForms.count > 0 Then
        lblTitle(10).FontBold = True
    End If

ElseIf glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes
    fraSalary.Height = 555
    fraSalary2.Top = 3310 '4850
    medNFacSalary.Visible = True
    lblNFacSal.Visible = True
ElseIf glbCompSerial = "S/N - 2429W" Then 'North Perth Ticket #19209 Franks 05/18/2011
    fraSalary.Height = 1150 '1515
    fraSalary2.Top = 3390
    Call ScreenNorthPerth
ElseIf glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #21097 Franks 11/02/2011
    'fraSalary.Height = 1150 '1515 'Ticket #22952
    'fraSalary2.Top = 3390 'Ticket #22952
    Call ScreenKNV
Else
    fraSalary.Height = 555
    'fraSalary.Width = 5150
    fraSalary2.Top = 2880 '4850
    
    'Re-arrange the tab sequence
    medPremium.TabStop = False
    medTotal.TabStop = False
    cboVGRoup.TabStop = False
    cboVStep.TabStop = False
End If
'end bryan

'added by bryan 24/Oct/05 Ticket#9607
If glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2396W" Then  'Aurora - Oshawa CHC Ticket #17341
    txtPayrollID.Enabled = False
End If

'WDGPHU - Ticket #17324
If glbCompSerial = "S/N - 2411W" Then
    lblPosGroup.Visible = True
    Call setCaption(lblPosGroup)
    txtPosGroup.Visible = True
End If

'Ticket #20652 - Town of Aurora
If glbCompSerial = "S/N - 2378W" Then
    lblPosGrp.Visible = True
    vbxTrueGrid.Columns(5).Visible = True
Else
    lblPosGrp.Visible = False
    vbxTrueGrid.Columns(5).Visible = False
End If

'Ticket #24482 - Town of Caledon - Using the VGroup field to store the Job's Division to create uniqueness between
'multiple same Position and Start Date positions linked to Salary.
If glbCompSerial = "S/N - 2182W" Then
    clpDiv.DataField = "SH_VGROUP"
    clpDiv.Visible = True
    lblDiv.Visible = True
    lblDiv.Caption = lStr("Division")
End If

'Release 8.1 - County of Wellington - Grey out Next Review Date
If glbCompSerial = "S/N - 2262W" Then
    lblTitle(10).Enabled = False
    dlpDate(1).Enabled = False
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

If glbCompSerial = "S/N - 2172W" Then 'Lanark Ticket #17221 by Frank 08/19/2009
    lblSalLevel.Top = lblGrid.Top
    lblSalLevel.Left = lblGrid.Left
    lblSalLevel.Visible = True
End If

glbOnTop = "FRMESALARY"

If glbSyndesis Then
    lblTitle(9).Caption = "Range"
    comSalScale.Tag = "00-Posion Grid Ranges"
End If

Call DecSetup

Call FIND_JOB

'Ticket #23499 - Reset label first to avoid conflict with Comments screen's menu label
lblComment.Caption = "SComments"

Call setCaption(lblTitle(12))
'Call setCaption(lblGrid)
lblGrid.Caption = lStr("Grid Category")
Call setCaption(lblComment)

'Ticket #23537 and Release 8.0
lblTitle(14).Caption = lStr("Hours/Week")

'Release 8.0 - Ticket #2268: Add Payroll ID to Label Master
lblPayID.Caption = lStr("Payroll ID")

'Jerry does not like "SComments" used from Label Master as original label so manually correcting it
If lblComment.Caption = "SComments" Then lblComment.Caption = "Comments"

comPayPer.Clear
comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour "
comPayPer.AddItem "Monthly "
'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
'If glbCompSerial = "S/N - 2282W" Then  'Ticket #17654
    comPayPer.AddItem "Daily "
'End If

'7.9 - Show Compa-Ration - Company Pref. setting
lblTitle(11).Visible = gsCompaRatio
lblCompaNum.Visible = gsCompaRatio

Call TabOrderSetup

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    
    'Ticket #20078 Franks 04/27/2011
    If glbCompSerial = "S/N - 2382W" Then  'Samuel
        Call getNONsecuritiesAgain(glbLEE_ID)
    End If
    
    If glbNoNONE Then
        If glbUNION = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then       'Hemu -EXE
        If glbUNION = "EXEC" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            If Len(glbBand) > 0 Then
                If InStr(1, ",A,B,C,D,E,", "," & glbBand & ",") = 0 Then
                    MsgBox "You Do Not Have Authority For This Transaction"
                    glbOnTop = Empty
                    Unload Me
                    Screen.MousePointer = DEFAULT
                    Exit Sub
                End If
            End If
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
    If glbNoEXEC Then
        If glbUNIONTe = "EXEC" Then     'Hemu -EXE
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            If Len(glbBand) > 0 Then
                If InStr(1, ",A,B,C,D,E,", "," & glbBand & ",") = 0 Then
                    MsgBox "You Do Not Have Authority For This Transaction"
                    glbOnTop = Empty
                    Unload Me
                    Screen.MousePointer = DEFAULT
                    Exit Sub
                End If
            End If
        End If
    End If
End If


If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then Exit Sub

If glbWFC Then
    Call Set_COMPA
    Call fgetSection(lblEEID.Caption)
    'If fSection = "GREN" Then
        lblTitle(12).FontBold = True
    'End If
End If
Screen.MousePointer = HOURGLASS

Call DoWFCGrids(False)

If glbCompSerial = "S/N - 2291W" Then 'syndesis
    lblBANDCode.DataField = ""
    lblBand.Caption = "Mid-Point"
    lblBand.Visible = True
    lblBANDCode.Visible = True
End If

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = IIf(glbSetSal, "Set ", "") & "Salary History- " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

If glbPayWeb Or glbVadim Or glbLambton Or glbInsync Or glbCompSerial = "S/N - 2351W" Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2380W" Then
    If Not glbCompSerial = "S/N - 2386W" And Not glbCompSerial = "S/N - 2382W" Then   'The Walter Fedy Partnership Ticket #14003 'Samuel
        lblTitle(12).FontBold = True
    End If
End If

If glbCompSerial = "S/N - 2382W" Then 'Samuel Ticket #20648 Franks 09/26/2011
    chkRedCircled.DataField = "SH_RED_CIRCLED"
    chkRedCircled.Visible = True
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" And glbMulti Then
    'Show Primary Position checkbox
    chkPrimary.Visible = True
    chkPrimary.Top = chkRedCircled.Top
    chkPrimary.Left = chkRedCircled.Left
Else
    'Hide Primary Position checkbox
    chkPrimary.Visible = False
End If

lblEENum.Caption = ShowEmpnbr(lblEEID)

lblEEID = glbLEE_ID

Call CR_JobHis_Snap(False)  ''Ticket #21511 - added the parameter
Call Set_Position(fglbJob$, False)

clpGrid.TABLTitle = lStr(lblGrid)

Call Display_Value

If glbCompSerial = "S/N - 2191W" Then
    fglbFrmt = "0.0"
    lblTitle(12).Caption = "Pay Type"
    clpCode(4).TABLTitle = "Pay Type Codes"
    clpCode(4).Tag = "Enter Pay Type Code"
Else
    fglbFrmt = "00"
End If

If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2425W" Then 'Ticket #16478 Samuel, Ticket #18221 - Four Villages CHC
    'Ticket #21988 Franks 05/02/2012 - comment out the following code, use Label Master
    'lblTitle(12).Caption = "Pay Type"
    'clpCode(4).TABLTitle = "Pay Type Codes"
    'clpCode(4).Tag = "Enter Pay Type Code"
    lblTitle(12).FontBold = True
End If
If glbCompSerial = "S/N - 2397W" Then  'Red Door Family Shelter Ticket #15255
    lblTitle(12).Caption = "Client Code"
    clpCode(4).TABLTitle = "CLIENT CODE"
    clpCode(4).Tag = "Enter Client Code"
End If
If glbCompSerial = "S/N - 2383W" Then 'Town of Orangeville Ticket #18844 Franks 01/13/2011
    lblTitle(12).Caption = "Pay Type"
    lblTitle(12).FontBold = True 'Ticket #20850 Franks 09/08/2011
    clpCode(4).TABLTitle = "Pay Type Codes"
    clpCode(4).Tag = "Enter Pay Type Code"
End If

If glbOttawaCCAC Or glbCompSerial = "S/N - 2229W" Or glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2171W" Then  'Ottawa CCAC, Inscape 'Ticket #24504 SPC
    lblTitle(12).FontBold = True
End If

If (glbCompSerial = "S/N - 2409W") Then lblTitle(12).FontBold = True 'Ticket #30066 Franks - Skylark Children

If glbCompSerial = "S/N - 2390W" Then  'Collectcorp Ticket #14889
    If glbEmpCountry = "U.S.A." Then
        lblTitle(12).FontBold = True
    End If
End If

If (glbCompSerial = "S/N - 2242W") Then 'C.C.A.C. London & Middlesex - Ticket #6718
    lblTitle(12).FontBold = True
    lblTitle(12).Caption = "Client #"
End If

If (glbCompSerial = "S/N - 2387W") Then 'Bird Packaging Limited  - TTicket #13166
    lblTitle(12).FontBold = True
    lblTitle(12).Caption = "Company Code"
End If

If glbCompSerial = "S/N - 2359W" Then 'Barber-Collins Security Services Ltd
    lblComment.Caption = lStr("Union")
    txtComment.Visible = False
    clpCode(5).Left = 1550 '1440
    clpCode(5).Top = 3630 '5600
    clpCode(5).Visible = True
End If

clpGrid.TextBoxWidth = 1000

'Ticket #15546
If glbLinamar Then
    lblTitle(10).FontBold = True
End If

If glbCompSerial = "S/N - 2373W" Then
    Call ScreenDMuskoka
End If


Call INI_Controls(Me)
clpGrid.SecurityMaintainable = False
Screen.MousePointer = DEFAULT
Exit Sub

Err_Deal:
If Err = 364 Then Resume Next

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
On Error GoTo Eh
Dim c As Long

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    panWindow.Height = Me.ScaleHeight - (panEEDesc.Height + vbxTrueGrid.Height + panControls.Height + 200)
    panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
    If panWindow.Height >= 7000 Then   '+ 230 Then
        scrControl.Value = 0
        panDetails.Top = 0
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = panWindow.Height
    End If

End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OCC_HEALTH_SAFETY", "edit/Add")
    Resume exH

End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmESALARY = Nothing
    Call NextForm
End Sub

Private Sub GetPlantCode()
Dim SQLQ As String, xPlantCode
Dim rsXEMP As New ADODB.Recordset
    glbEmpPlant = ""
    locCountry = ""
    SQLQ = "SELECT ED_EMPNBR,ED_SECTION,ED_COUNTRY FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsXEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsXEMP.EOF Then
        xPlantCode = rsXEMP("ED_SECTION")
        locCountry = rsXEMP("ED_COUNTRY")
    End If
    rsXEMP.Close
    glbEmpPlant = xPlantCode
End Sub
Private Function getJOB(nJob As String, nGrid As String)
Dim SQLQ As String, X%, xLev
Dim Msg$
Dim rsJOB As New ADODB.Recordset
Dim rsDESCR2 As New ADODB.Recordset
'Dim rsGrid As New ADODB.Recordset
'Dim xGridList

getJOB = False

On Error GoTo Jobd_Err

Call setGridList(nJob)

If Len(nJob) > 0 Then
    If glbMultiGrid Then
        SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE = '" & nJob$ & "' AND JB_GRID='" & nGrid & "'"
    Else
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & nJob$ & "' "
    End If
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsJOB.EOF Then
        fglbBAND = ""
        Exit Function
    End If
    
    If glbCompSerial = "S/N - 2291W" Then
        If Not IsNull(rsJOB("JB_MIDPOINT")) And rsJOB("JB_MIDPOINT") > 0 And rsJOB("JB_MIDPOINT") < 12 Then
            lblBANDCode.Caption = Format(rsJOB("JB_S" & rsJOB("JB_MIDPOINT")), "$0.00")
        Else
            lblBANDCode.Caption = "$0.00"
        End If
    End If
    
    If glbWFC Then fglbBAND = IIf(IsNull(rsJOB("JB_BAND")), "", rsJOB("JB_BAND"))
    
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For x% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        If Not IsNull(rsJOB("JB_S" & X%)) Then JobSnaps_PayScale(X) = Round2DEC(rsJOB("JB_S" & X%))
        
        If glbCompSerial = "S/N - 2378W" And rsJOB("JB_SALCD") <> lblSalCode Then      'Town of Aurora
            If Not IsNull(rsJOB("JB_S" & X% & "A")) Then JobSnaps_PayScale(X) = Round2DEC(rsJOB("JB_S" & X% & "A"))
        End If
    Next
    If Not IsNull(rsJOB("JB_SALCD")) Then JobSnaps_Salary_Code$ = rsJOB("JB_SALCD")
    If Not IsNull(rsJOB("JB_MIDPOINT")) Then JobSnap_MidPoint! = rsJOB("JB_MIDPOINT")
    If Not IsNull(rsJOB("JB_FTEHRS")) Then
        JobSnaps_Salary_FTEHrs = rsJOB("JB_FTEHRS")
    Else
        JobSnaps_Salary_FTEHrs = 1
    End If
    fglbGridType = 0
    
    SQLQ = "SELECT JB_DESCR2,JB_ID FROM HRJOB WHERE JB_CODE='" & nJob & "'"
    rsDESCR2.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsDESCR2.EOF Then
        If IsNumeric(rsDESCR2("JB_DESCR2")) Then
            If Val(rsDESCR2("JB_DESCR2")) = 0.5 Then
                fglbGridType = 0.5
            End If
        End If
    End If
    rsDESCR2.Close
    comSalScale.Clear
    
    comSalScale.AddItem Format(0, fglbFrmt)
    
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'For x% = 1 To 11
    'For X% = 1 To 15
    For X% = 1 To 20
        If rsJOB("jb_s" & Trim(Str(X%))) <> 0 Then
            xLev = X%
            If fglbGridType = 0.5 Then xLev = (X% + 1) / 2
            
            If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
                If xLev = 1 Then
                    comSalScale.AddItem "Start"
                Else
                    comSalScale.AddItem Format(xLev - 1, fglbFrmt)
                End If
            Else
                comSalScale.AddItem Format(xLev, fglbFrmt)
            End If
        End If
    Next
    
    If fglbGridType = 0.5 And Val(lblSalaryGrade) <> 0 Then
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
            End If
        Else
            comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
        End If
    Else
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
            End If
        Else
            comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
        End If
    End If
    
    If glbWFC Then
        Call Set_MarketLine_List
    End If
End If

getJOB = True

Exit Function

Jobd_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "HRJOB", "SELECT")
Resume Next

End Function

Sub Set_MarketLine_List()
Dim rsWFC As New ADODB.Recordset
Dim X%, I%
Dim xItemAdd
Dim SQLQ

SQLQ = "select MarketLine from WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & lblBANDCode & "'"
If Len(clpCode(0)) > 0 Then
    SQLQ = SQLQ & " AND SectionCode ='" & clpCode(0) & "'"
End If
If Len(txtFiscalYear) > 0 Then
    SQLQ = SQLQ & " AND FiscalYear =" & txtFiscalYear & ""
End If
SQLQ = SQLQ & " group by MarketLine"

rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenKeyset
X% = 0
cmbMarketLine.Clear
Do Until rsWFC.EOF
    cmbMarketLine.AddItem rsWFC("marketline")
    If rsWFC("marketline") = txtMarketLine Then
        cmbMarketLine.ListIndex = X%
    End If
    X% = X% + 1
    rsWFC.MoveNext
Loop
rsWFC.Close
'MarketLine_Desc Me
Call SalMarketLineDesc

End Sub
Private Sub SalMarketLineDesc()
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    If Len(Trim(cmbMarketLine)) > 0 Then
        SQLQ = "SELECT TB_KEY,TB_DESC FROM HRTABL WHERE TB_NAME ='WFML' AND TB_KEY ='" & cmbMarketLine & "' "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            lblMLine.Caption = rsTemp("TB_DESC")
        End If
        rsTemp.Close
    End If
End Sub
Private Sub lblBANDCode_Change()
    Set_SalState
End Sub

Private Sub lblCompaNum_Change()
    If glbFrench Then
        If IsNumeric(lblCompaNum) Then
            lblCompaNum = Round(lblCompaNum, 2)
        End If
    Else
        lblCompaNum = Round(Val(lblCompaNum), 2)
    End If
End Sub

Private Sub lblSalaryGrade_Change()
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        lblSalaryGrade = Format(Val(lblSalaryGrade), "00")
    End If
    
    If fglbGridType = 0.5 And Val(lblSalaryGrade) > 0 Then
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
            End If
        Else
            comSalScale = Format((Val(lblSalaryGrade) + 1) / 2, fglbFrmt)
        End If
    Else
        If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
            If lblSalaryGrade = "01" Then
                comSalScale = "Start"
            Else
                comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
            End If
        Else
            comSalScale = Format(Val(lblSalaryGrade), fglbFrmt)
        End If
    End If
End Sub

Private Sub lblSalCode_Change()
If flagFrmLoad = False Then Exit Sub 'carmen may 00
If Len(lblSalCode.Caption) > 0 Then
    If lblSalCode.Caption = "A" Then
        comPayPer.ListIndex = 0
    ElseIf lblSalCode.Caption = "H" Then
        comPayPer.ListIndex = 1
    'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
    'ElseIf lblSalCode.Caption = "D" And glbCompSerial = "S/N - 2282W" Then
    ElseIf lblSalCode.Caption = "D" Then  'Ticket #17654
        comPayPer.ListIndex = 3
    Else
        comPayPer.ListIndex = 2
    End If
End If
End Sub
Sub Set_WFC_COMPA()
Dim xDollear
If glbWFC And UnionExecNone Then
    lblCompaNum = 0
    'If optUserSys(0) Then xDollear = Val(lblsalstate(1)) Else xDollear = Val(mskCampa)
    xDollear = Val(lblsalstate(1))
    'Changed by Bryan 22/Sep/05 Ticket#9343
    
    If Val(xDollear) <> 0 Then
        If glbCompSerial = "S/N - 2373W" Then
            lblCompaNum = (Val(medTotal) / xDollear) * 100
        Else
            lblCompaNum = (Val(medsalary) / xDollear) * 100
        End If
    End If
    If Val(lblCompaNum) > 999.99 Then lblCompaNum = "999.99"
    lblCompaNum.Caption = Format(lblCompaNum, "0.00")
End If
End Sub

Sub Set_SalaryLevel(xJobCode) 'County of Lanark
Dim rsTJob As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT JB_GRPCD FROM HRJOB WHERE JB_CODE = '" & xJobCode & "' "
    rsTJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTJob.EOF Then
        If Not IsNull(rsTJob("JB_GRPCD")) Then
            lblSalLevel.Caption = "Salary Level: " & GetTABLDesc("JBGC", rsTJob("JB_GRPCD"))
        Else
            lblSalLevel.Caption = "Salary Level: "
        End If
    End If
    rsTJob.Close
    
End Sub

Sub Set_SalState()
Dim SQLQ
Dim rsWFC As New ADODB.Recordset
Dim xPlantCd
If Not glbWFC Then Exit Sub
xPlantCd = glbEmpPlant
If Len(clpCode(0).Text) > 0 Then
    xPlantCd = clpCode(0).Text
End If
SQLQ = "SELECT LDOLLARS,MDOLLARS,HDOLLARS FROM WFC_Salary_Administration "
SQLQ = SQLQ & " WHERE [BAND]='" & Trim(lblBANDCode) & "'"
SQLQ = SQLQ & " AND [MARKETLINE]='" & IIf(txtMarketLine.Visible, txtMarketLine, cmbMarketLine) & "'"
SQLQ = SQLQ & " AND SectionCode='" & xPlantCd & "' "
If Len(txtFiscalYear) > 0 Then
    If IsNumeric(txtFiscalYear) Then
        SQLQ = SQLQ & " AND FiscalYear='" & txtFiscalYear & "' "
    End If
End If

rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenStatic

If rsWFC.EOF Then
  lblsalstate(0) = "": lblsalstate(1) = "": lblsalstate(2) = ""
Else
  lblsalstate(0) = Format(rsWFC("LDOLLARS"), "0.00")
  lblsalstate(1) = Format(rsWFC("MDOLLARS"), "0.00")
  lblsalstate(2) = Format(rsWFC("HDOLLARS"), "0.00")
End If
rsWFC.Close
End Sub


Private Sub medAmtChng_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
'Hemu - essex
'fglbAmtOld(Index) = CCur(Val(medAmtChng(Index)))  'Jaddy 10/25/99
'Hemu - essex
End Sub


Private Sub medAmtChng_KeyPress(Index As Integer, KeyAscii As Integer)
    ' dkostka - 01/12/01 - Fixed problem where salary would change if tabbing past step
    '   by disabling step if they have used any other salary-changing functions.
    comSalScale.Enabled = False
End Sub

Private Sub medAmtChng_LostFocus(Index As Integer)
If glbCompSerial = "S/N - 2436W" Then  'Family Day - Ticket #21152 Franks 04/01/2013
    Exit Sub
End If
If glbSetSal Then Exit Sub
If Not IsNumeric(medAmtChng(Index)) Then
   medAmtChng(Index) = 0
End If

If Not IsNumeric(fglbAmtOld(Index)) Then
   fglbAmtOld(Index) = 0
End If

If medAmtChng(Index) <> fglbAmtOld(Index) Then
    If medAmtChng(Index) <> 0 Then
        If glbFrench Then
            If orgSalary > 0 Then
                medPercentChng(Index) = medAmtChng(Index) / orgSalary
            Else
                medPercentChng(Index) = 1
            End If
        Else
            If Val(orgSalary) > 0 Then
                medPercentChng(Index) = medAmtChng(Index) / orgSalary
            Else
                medPercentChng(Index) = 1
            End If
        End If
    End If
    Call Upd_Salary
End If

Call PerOrSal

End Sub

Private Sub medNFacSalary_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPercentChng_GotFocus(Index As Integer)

Call SetPanHelp(ActiveControl)

If medPercentChng(Index) = "" Then
   medPercentChng(Index) = 0
End If

medPercentChng(Index) = medPercentChng(Index) * 100
fglbPCOld(Index) = medPercentChng(Index)

End Sub

Private Sub medPercentChng_KeyPress(Index As Integer, KeyAscii As Integer)
    ' dkostka - 01/12/01 - Fixed problem where salary would change if tabbing past step
    '   by disabling step if they have used any other salary-changing functions.
    comSalScale.Enabled = False
End Sub

Private Sub medPercentChng_LostFocus(Index As Integer)
If glbCompSerial = "S/N - 2436W" Then  'Family Day - Ticket #21152 Franks 04/01/2013
    Exit Sub
End If
If Not IsNumeric(medPercentChng(Index)) Then
   medPercentChng(Index) = 0
End If

If Not IsNumeric(fglbPCOld(Index)) Then
   fglbPCOld(Index) = 0
End If
If Not glbSetSal Then
    If medPercentChng(Index) <> fglbPCOld(Index) Then
        ' DK - 03/16/2000 - Removed encryption code
        ' -----
        medAmtChng(Index) = CDbl(medPercentChng(Index)) * orgSalary / 100
        ' -----
        Call Upd_Salary
    End If
End If

medPercentChng(Index) = medPercentChng(Index) / 100
If Not glbSetSal Then
    Call PerOrSal
End If
End Sub

Private Sub medPremium_Change()
Call setPayPeriodSalary
End Sub

Private Sub medPremium_LostFocus()
Dim X%

On Error GoTo Salary_Err 'uncommented 28July99
If Not IsNumeric(medsalary) Then medsalary = 0
If glbFrench Then
    medsalary = Round2DEC(medsalary)    'Val() causing the values to trunc to 0 decimal places
Else
    medsalary = Round2DEC(Val(medsalary))
End If

If Not IsNumeric(medPremium) Then medPremium = 0
If glbFrench Then
    medPremium = Round2DEC(medPremium)  'Val() causing the values to trunc to 0 decimal places
Else
    medPremium = Round2DEC(Val(medPremium))
End If

'DMuskoka
If glbCompSerial = "S/N - 2373W" Then
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
    'Ticket #27106 - They want to compute the Step # using the Salary Amount and not Total
    'Call Set_SalaryGrade(Val(medTotal))
    Call Set_SalaryGrade(Val(medsalary))
Else
    If glbFrench Then
        Call Set_SalaryGrade_French(medsalary)  'Val() causing the values to trunc to 0 decimal places
    Else
        Call Set_SalaryGrade(Val(medsalary))
    End If
End If
Exit Sub

Salary_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "medPremium", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub

Private Sub medSalary_Change()
    Call setPayPeriodSalary
End Sub

Sub setPayPeriodSalary()
    If glbCompSerial = "S/N - 2373W" Then 'muskoka
        If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
            medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
        ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
            medTotal.Text = medsalary.Text
        End If
        If IsNumeric(medTotal) Then
            'Hemu - 08/11/2003 Begin - Calculate and Display Salary per Pay Period
            If fglbPhrs <> 0 Then
                If lblSalCode = "H" Then
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                    lblPayPeriodSalary = Round2DEC(Val(medTotal) * fglbPhrs)
                    lblHoursPay.Visible = False
                    lblTitle(21).Visible = False
                ElseIf lblSalCode = "M" Then
                    'lblPayPeriodSalary = Round2DEC(Val(medTotal))
                    lblPayPeriodSalary.Visible = False
                    lblTitle(15).Visible = True
                    'lblPayPeriodSalary = (Round2DEC(Val(medTotal) / (fglbPhrs * 2)) * fglbPhrs)
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    lblHoursPay = Round2DEC(Val(medTotal) / (fglbPhrs))
                'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
                ElseIf lblSalCode = "D" Then
                    If IsDate(dlpDate(0)) Then
                        lblPayPeriodSalary.Visible = True
                        lblTitle(15).Visible = True
                    
                        If GetLeapYear(Year(dlpDate(0))) Then
                            lblPayPeriodSalary = Round2DEC(((Val(medTotal) / 366) / fglbWhrs#) * fglbPhrs)
                        Else
                            lblPayPeriodSalary = Round2DEC(((Val(medTotal) / 365) / fglbWhrs#) * fglbPhrs)
                        End If
                        
                        'Ticket #17654 - Opening up for everyone - the correct formula:
                        'Ticket #21082 Franks 10/18/2011
                        If fglbDhrs = 0 Then
                            lblPayPeriodSalary = 0
                        Else
                            lblPayPeriodSalary = Round2DEC((Val(medTotal) / fglbDhrs) * fglbPhrs)
                        End If
                    End If
                ElseIf fglbWhrs# = 0 Then
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                
                    lblPayPeriodSalary = 0
                    lblHoursPay = 0
                Else
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                
                    lblPayPeriodSalary = Round2DEC(((Val(medTotal) / 52) / fglbWhrs#) * fglbPhrs)
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    
                    'City of Niagara Falls - Special Hourly Rate calculation
                    If glbCompSerial = "S/N - 2276W" Then
                        If fglbDhrs = 0 Or fglbNiagPhrs = 0 Then 'Ticket #14175
                            lblHoursPay = 0
                            lblPayPeriodSalary = 0
                        Else
                            'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                            'So xPHrs contains Pay Periods per Year (SH_PAYP) and xWHRS contains Hours Per Pay (JB_DHRS)
                            'lblHoursPay = Round2DEC((Val(medTotal) / fglbNiagPhrs) / (fglbDhrs * 5))
                            lblHoursPay = Round2DEC((Val(medTotal) / fglbNiagPhrs / fglbDhrs))
                            
                            'Ticket #24559 - Pay Period Amount not getting recomputed when Pay Period changes
                            lblPayPeriodSalary = Round((Val(medsalary) / fglbNiagPhrs), 4)
                        End If
                    Else
                        lblHoursPay = Round2DEC((Val(medTotal) / 52) / fglbWhrs#)
                    End If
                End If
                If lblPayPeriodSalary.Visible = True Then
                    lblPayPeriodSalary = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                End If
                If lblSalCode <> "H" Then
                    lblHoursPay = Format(lblHoursPay, "#0." & String(glbCompDecHR, "0"))
                End If
            Else
                lblPayPeriodSalary.Visible = True
                lblTitle(15).Visible = True
                
                lblPayPeriodSalary = 0
                lblHoursPay = 0
            End If
            'Hemu - 08/11/2003 End
        Else
            lblPayPeriodSalary = 0
            lblHoursPay = 0
        End If
    Else
        If IsNumeric(medsalary) Then
            'Hemu - 08/11/2003 Begin - Calculate and Display Salary per Pay Period
            If fglbPhrs <> 0 Then
                lblTitle(21).Caption = "Hourly Rate :"
                If lblSalCode = "H" Then
                    lblTitle(15).Caption = "Salary Per Pay : "
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                    
                    If glbCompSerial = "S/N - 2344W" Then 'cascade
                        'Ticket #29233 - Change the formula
                        'lblPayPeriodSalary = Round2DEC(medsalary)
                        lblPayPeriodSalary = Round2DEC(medsalary * fglbPhrs) '86.67)
                    ElseIf glbWFC Then 'Ticket #13520
                        If clpCode(4).Text = "D" Then 'Pay Period = D
                            lblPayPeriodSalary = Round2DEC(medsalary * 5)
                        Else 'Pay Period = W
                            lblPayPeriodSalary = Round2DEC(medsalary * fglbPhrs)
                        End If
                    ElseIf glbCompSerial = "S/N - 2407W" Then 'Farmer Mutual - Mostafa changed as per jerry's request
                            'Ticket #24948 - Changed the logic, i.e.
                                'SalaryPerDay = Annual Salary / 261
                                'SalaryPerHour = SalaryPerDay / HoursPerDay
                                'SalaryPerPay = SalaryPerHours * HoursPerPay
                            'lblHoursPay = Round((medsalary / 261) / fglbDhrs, 2)
                            lblPayPeriodSalary = Round(medsalary * fglbPhrs, 2)
                    Else
                        lblPayPeriodSalary = Round2DEC(medsalary * fglbPhrs)
                    End If
                    
                    lblHoursPay.Visible = False
                    
                    'lblTitle(21).Visible = False
                    'Ticket #20396 - Jerry asked to show the Annual Salary for everyone
                    'If glbCompSerial = "S/N - 2172W" Then 'County of Lanark Ticket #17076
                        If glbCompSerial = "S/N - 2407W" Then 'Farmer Mutual - Mostafa changed as per jerry's request
                            lblHoursPay.Caption = Format(Round(medsalary * fglbDhrs * 261, 2), "#0." & String(2, "0"))
                        Else
                            If glbCompSerial = "S/N - 2344W" Then 'Cascade - Ticket #25449
                                lblHoursPay.Caption = Round2DEC(medsalary * fglbPhrs * 24)
                            Else
                                lblHoursPay.Caption = Round2DEC(medsalary * fglbWhrs# * 52)
                            End If
                            lblHoursPay.Caption = Format(lblHoursPay, "#0." & String(glbCompDecHR, "0"))
                        End If
                        lblHoursPay.Visible = True
                        lblTitle(21).Caption = "Annual Salary : "
                        lblTitle(21).Visible = True
                    'Else
                    '    lblTitle(21).Visible = False
                    'End If
                    
                ElseIf lblSalCode = "M" Then
                    'lblPayPeriodSalary = Round2DEC(medsalary)
                    'lblPayPeriodSalary = (Round2DEC(Val(medsalary) / (fglbPhrs * 2)) * fglbPhrs)
                    
                    'Ticket #20396 - Jerry asked to show the Annual Salary for everyone
                    'lblPayPeriodSalary.Visible = False
                    'lblTitle(15).Visible = False
                    lblPayPeriodSalary.Caption = Round2DEC(medsalary * 12)
                    lblPayPeriodSalary.Caption = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Caption = "Annual Salary : "
                    lblTitle(15).Visible = True
                    
                    If fglbWhrs# <> 0 Then
                        lblHoursPay.Visible = True
                        lblTitle(21).Visible = True
                        If glbFrench Then
                            'Ticket #20396 - Jerry said this formula is not correct so gave the correct one
                            'lblHoursPay = Round2DEC(medsalary / (fglbPhrs))
                            lblHoursPay = Round2DEC(((medsalary * 12) / 52) / (fglbWhrs#))
                        Else
                            'Ticket #20396 - Jerry said this formula is not correct so gave the correct one
                            'lblHoursPay = Round2DEC(Val(medsalary) / (fglbPhrs))
                            lblHoursPay = Round2DEC(((Val(medsalary) * 12) / 52) / (fglbWhrs#))
                        End If
                    Else
                        lblHoursPay.Visible = False
                        lblTitle(21).Visible = False
                    End If
                'woodbridge get's Daily salary - Bryan 19/Sep/05 Ticket #9354
                ElseIf lblSalCode = "D" Then
                    'Ticket #20396 - Jerry asked not to show the Salary per Pay instead show Annual Sal
                    'If IsDate(dlpDate(0)) Then
                        'Ticket #20396 - Jerry asked to show the Annual Salary for everyone
                        lblPayPeriodSalary.Visible = True
                        lblTitle(15).Visible = True
                        lblPayPeriodSalary.Caption = Round2DEC((medsalary * 5) * 52) 'Assumes 5 day/week
                        lblPayPeriodSalary.Caption = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                        lblTitle(15).Caption = "Annual Salary : "

                        'Ticket #20396 - Jerry asked not to show the Salary per Pay instead Annual Sal is shown above
                        'If GetLeapYear(Year(dlpDate(0))) Then
                        '    lblPayPeriodSalary = Round2DEC(((medsalary / 366) / fglbWhrs#) * fglbPhrs)
                        'Else
                        '    lblPayPeriodSalary = Round2DEC(((medsalary / 365) / fglbWhrs#) * fglbPhrs)
                        'End If
                        
                        'Ticket #20396 - Jerry asked not to show the Salary per Pay instead Annual Salary is shown above
                        'Ticket #17654 - Opening up for everyone - the correct formula:
                        'lblPayPeriodSalary = Round2DEC((medsalary / fglbDhrs) * fglbPhrs)
                        
                        lblHoursPay.Visible = True
                        lblTitle(21).Visible = True
                        If fglbDhrs <> 0 Then
                            If glbFrench Then
                                lblHoursPay = Round2DEC(medsalary / fglbDhrs)
                            Else
                                lblHoursPay = Round2DEC(Val(medsalary) / fglbDhrs)
                            End If
                        Else
                            lblHoursPay = 0
                        End If
                    'End If
                ElseIf fglbWhrs# = 0 Then
                    lblTitle(15).Caption = "Salary Per Pay : "
                    
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                
                    lblPayPeriodSalary = 0
                    lblHoursPay = 0
                Else
                    If fglbPhrs = "" Then fglbPhrs = 1
                    lblTitle(15).Caption = "Salary Per Pay : "
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                    
                    If glbCompSerial = "S/N - 2407W" Then 'Farmer Mutual - Mostafa changed as per jerry's request
                        'Ticket #24948 - Changed the logic, i.e.
                            'SalaryPerDay = Annual Salary / 261
                            'SalaryPerHour = SalaryPerDay / HoursPerDay
                            'SalaryPerPay = SalaryPerHours * HoursPerPay
                        'lblPayPeriodSalary = Round2DEC(medsalary / 26.07143)
                        lblHoursPay = Round((medsalary / 261) / fglbDhrs, 2)
                        lblPayPeriodSalary = Round(((medsalary / 261) / fglbDhrs) * fglbPhrs, 2)
                    ElseIf glbCompSerial = "S/N - 2174W" Then  'Ticket #23382 Franks 02/27/2014
                        lblPayPeriodSalary = Round((medsalary / 24), 2)
                    Else
                        lblPayPeriodSalary = Round2DEC(((medsalary / 52) / fglbWhrs#) * fglbPhrs)
                    End If
                    
                    'lblPayPeriodSalary = Round2DEC(((medsalary / 52) / fglbWhrs#) * fglbPhrs)
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                    
                    'City of Niagara Falls - Special Hourly Rate calculation
                    If glbCompSerial = "S/N - 2276W" Then
                        If fglbDhrs = 0 Or fglbNiagPhrs = 0 Then 'Ticket #14175
                            lblHoursPay = 0
                            lblPayPeriodSalary = 0
                        Else
                            'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                            'So xPHrs contains Pay Periods per Year (SH_PAYP) and xWHRS contains Hours Per Pay (JB_DHRS)
                            'lblHoursPay = Round2DEC((Val(medsalary) / fglbNiagPhrs) / (fglbDhrs * 5))
                            lblHoursPay = Round2DEC((Val(medsalary) / fglbNiagPhrs / fglbDhrs))
                            
                            'Ticket #24559 - Pay Period Amount not getting recomputed when Pay Period changes
                            lblPayPeriodSalary = Round((Val(medsalary) / fglbNiagPhrs), 4)
                        End If
                    ElseIf glbCompSerial = "S/N - 2344W" Then 'cascade Ticket #12293
                        'Ticket #29233 - Change the formula
                        'lblPayPeriodSalary = Round2DEC((Val(medsalary) / 24))
                        lblPayPeriodSalary = Round2DEC((Val(medsalary) / 24 / 86.67) * fglbPhrs) '86.67)
                        lblHoursPay = 0
                        lblHoursPay.Visible = False
                        lblTitle(21).Visible = False
                    Else
                        'Ticket #13045
                        'lblHoursPay = Round2DEC((Val(medsalary) / 52) / fglbWhrs#)
                        If Val(txtWHRS.Text) <> 0 Then 'Ticket #13210
                            If glbFrench Then
                                lblHoursPay = Round2DEC((medsalary / 52) / Val(txtWHRS.Text))
                            Else
                                lblHoursPay = Round2DEC((Val(medsalary) / 52) / Val(txtWHRS.Text))
                            End If
                        Else
                            lblHoursPay = 0
                        End If
                    End If
                    
                    If glbCompSerial = "S/N - 2390W" Then  'Gatestone - Ticket #22122
                        lblPayPeriodSalary = Round2DEC((Val(medsalary) / 24))
                    End If
                    
                    If glbCompSerial = "S/N - 2436W" Then  'Family Day Care Services - Ticket #22316
                        lblPayPeriodSalary = Round2DEC((Val(medsalary) / 24))
                    End If
                    
                    If glbCompSerial = "S/N - 2457W" Then  'McLeod Law - Ticket #24863
                        lblPayPeriodSalary = Round2DEC((Val(medsalary) / 24))
                    End If
                    
                    If glbCompSerial = "S/N - 2407W" Then 'Farmer Mutual - Mostafa changed as per jerry's request
                        'Ticket #24948 - Changed the logic, i.e.
                            'SalaryPerDay = Annual Salary / 261
                            'SalaryPerHour = SalaryPerDay / HoursPerDay
                            'SalaryPerPay = SalaryPerHours * HoursPerPay
                        lblHoursPay = Round((medsalary / 261) / fglbDhrs, 2)
                        lblPayPeriodSalary = Round(((medsalary / 261) / fglbDhrs) * fglbPhrs, 2)
                    End If
                    
                    If glbWFC And lblSalCode = "A" Then 'Ticket #13520 For Annual Salary
                        'If glbWFC And fSection = "GREN" Then
                        '    If clpCode(4).Text = "M" Then
                        '        lblPayPeriodSalary = Round2DEC(medsalary / 12)
                        '    End If
                        'End If
                        If clpCode(4).Text = "SM" Then
                            lblPayPeriodSalary = Round2DEC(medsalary / 24)
                        End If
                        If clpCode(4).Text = "M" Then
                            lblPayPeriodSalary = Round2DEC(medsalary / 12)
                        End If
                        If clpCode(4).Text = "W" Then
                            lblPayPeriodSalary = Round2DEC(medsalary / 52)
                        End If
                    End If
                End If
                'If glbWFC And fSection = "GREN" Then
                '    If clpCode(4).Text = "M" Then
                '        lblPayPeriodSalary = Round2DEC(medsalary / 12)
                '    End If
                'End If

                If glbCompSerial = "S/N - 2382W" Then 'Namasco
                    If lblSalCode <> "H" Then
                        lblPayPeriodSalary.Visible = True
                        lblTitle(15).Visible = True
                    
                        lblPayPeriodSalary = Round2DEC(medsalary / 24)
                    End If
                End If
                
                If lblPayPeriodSalary.Visible = True Then
                    lblPayPeriodSalary = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                End If
                
                If lblSalCode <> "H" Then 'Houlry Rate
                    If glbWFC And lblSalCode = "A" Then 'Ticket #13520 For Annual Salary
                        If clpCode(4).Text = "SM" Then
                            lblHoursPay = Round2DEC((medsalary / 24) / fglbPhrs, "Y")
                        End If
                        If clpCode(4).Text = "M" Then
                            lblHoursPay = Round2DEC((medsalary / 12) / fglbPhrs, "Y")
                        End If
                        If clpCode(4).Text = "W" Then
                            lblHoursPay = Round2DEC((medsalary / 52) / fglbPhrs, "Y")
                        End If
                    Else
                        lblHoursPay = Format(lblHoursPay, "#0." & String(glbCompDecHR, "0"))
                    End If
                    
                    If glbCompSerial = "S/N - 2382W" Then 'Samuel
                        'lblHoursPay = Format(lblPayPeriodSalary / 86.66, "#0." & String(glbCompDecHR, "0"))
                        'Ticket #24550 Franks 10/25/2013 - use fglbPhrs to calculate
                        If fglbPhrs = 0 Then
                            lblHoursPay = Format(0, "#0." & String(glbCompDecHR, "0"))
                        Else
                            lblHoursPay = Format(lblPayPeriodSalary / fglbPhrs, "#0." & String(glbCompDecHR, "0"))
                        End If
                    End If
                End If
            Else
                'City of Niagara Falls - Special Hourly Rate calculation
                If glbCompSerial = "S/N - 2276W" Then
                    lblTitle(15).Caption = "Salary Per Pay : "
                    lblPayPeriodSalary.Visible = True
                    lblTitle(15).Visible = True
                    
                    lblHoursPay.Visible = True
                    lblTitle(21).Visible = True
                
                    If fglbDhrs = 0 Or fglbNiagPhrs = 0 Then 'Ticket #14175
                        lblHoursPay = 0
                        lblPayPeriodSalary = 0
                    Else
                        'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                        'So xPHrs contains Pay Periods per Year (SH_PAYP) and xWHRS contains Hours Per Pay (JB_DHRS)
                        'lblHoursPay = Round2DEC((Val(medsalary) / fglbNiagPhrs) / (fglbDhrs * 5))
                        lblHoursPay = Round2DEC((Val(medsalary) / fglbNiagPhrs / fglbDhrs))
                        
                        'Ticket #24559 - Pay Period Amount not getting recomputed when Pay Period changes
                        lblPayPeriodSalary = Round((Val(medsalary) / fglbNiagPhrs), 4)
                    
                        If lblPayPeriodSalary.Visible = True Then
                            lblPayPeriodSalary = Format(lblPayPeriodSalary, "#0." & String(glbCompDecHR, "0"))
                        End If
                        
                        If lblSalCode <> "H" Then 'Houlry Rate
                            lblHoursPay = Format(lblHoursPay, "#0." & String(glbCompDecHR, "0"))
                        End If
                    End If
                Else
                    lblPayPeriodSalary = 0
                    lblHoursPay = 0
                End If
            End If
        Else
            lblPayPeriodSalary = 0
            lblHoursPay = 0
            'Hemu - 08/11/2003 End
        End If
    End If
End Sub

Private Sub medSalary_GotFocus()

Call SetPanHelp(ActiveControl)
If glbFrench Then
    If IsNumeric(medsalary) Then
        fglbSHold@ = CCur(medsalary)
    End If
Else
    fglbSHold@ = CCur(Val(medsalary))
End If

End Sub

Private Sub medSalary_KeyPress(KeyAscii As Integer)
    ' dkostka - 01/12/01 - Fixed problem where salary would change if tabbing past step
    '   by disabling step if they have used any other salary-changing functions.
    comSalScale.Enabled = False
End Sub

Private Sub medSalary_LostFocus()
Dim X%

On Error GoTo Salary_Err 'uncommented 28July99
If Not IsNumeric(medsalary) Then medsalary = 0
If glbFrench Then
    medsalary = Round2DEC(medsalary)    'Val() causing the values to trunc to 0 decimal places
Else
    medsalary = Round2DEC(Val(medsalary))
End If

If Not IsNumeric(medPremium) Then medPremium = 0
If glbFrench Then
    medPremium = Round2DEC(medPremium)  'Val() causing the values to trunc to 0 decimal places
Else
    medPremium = Round2DEC(Val(medPremium))
End If

If Not glbSetSal Then
    Call setPercent
End If

'DMuskoka
If glbCompSerial = "S/N - 2373W" Then
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
    
    'Ticket #26837 - Do not reset the Salary Grade/Step when the Salary is not Current or New
    If fglbNew Or chkCurrent Then
        'Ticket #27106 - They want to compute the Step # using the Salary Amount and not Total
        'Call Set_SalaryGrade(Val(medTotal))
        Call Set_SalaryGrade(Val(medsalary))
    End If
Else
    If glbFrench Then
        'Ticket #26837 - Do not reset the Salary Grade/Step when the Salary is not Current or New
        If fglbNew Or chkCurrent Then
            Call Set_SalaryGrade_French(medsalary)  'Val() causing the values to trunc to 0 decimal places
        End If
    Else
        'Ticket #26837 - Do not reset the Salary Grade/Step when the Salary is not Current or New
        If fglbNew Or chkCurrent Then
            Call Set_SalaryGrade(Val(medsalary))
        End If
    End If
End If

Exit Sub

Salary_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "medsalary", "HR_SALARY_HISTORY", "Update")
Resume Next
Unload Me

End Sub


Private Sub PerOrSal()  'RAUBREY 6/6/97
'Ticket #20666 Franks 07/19/2011
'This caused problem when Jerry tested it with customer, disable it
'If glbFrench Then
'    If medAmtChng(1) = 0 And medAmtChng(2) = 0 And medAmtChng(3) = 0 Then
'        fraSalary.Enabled = True
'    Else
'        fraSalary.Enabled = False
'    End If
'Else
'    If Val(medAmtChng(1)) = 0 And Val(medAmtChng(2)) = 0 And Val(medAmtChng(3)) = 0 Then
'        fraSalary.Enabled = True
'    Else
'        fraSalary.Enabled = False
'    End If
'End If
End Sub


Private Function Round2DEC(tmpNUM, Optional HourlyRate As String)    'laura nov 10, 1997
Dim strNUM As String, X%

If glbFrench Then
    tmpNUM = Replace(Replace(tmpNUM, ",", "."), " ", "")
End If

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    If GetEmpData(glbLEE_ID, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(Val(tmpNUM), glbCompDecHR)
End If
If glbWFC And locCountry = "AUSTRALIA" Then
    If HourlyRate = "Y" Then
        locCompDecHR = 4
    End If
    Round2DEC = Round(tmpNUM, locCompDecHR)
End If
End Function

Private Function Set_Position(nJob As String, nCurrent As Boolean)
Dim SQLQ As String, Msg$
Dim rsHRJOB As New ADODB.Recordset

Set_Position = False

On Error GoTo SCError

Screen.MousePointer = HOURGLASS

dynaJobHIS.Requery
dynaJobHIS.Filter = ""
SQLQ = ""

If nCurrent Then SQLQ = SQLQ & " JH_CURRENT<>0 "
If nJob <> "" Then
    'Ticket #29233 - Get the right Job information based on Start Date as well - esp. when selecting older Salary record
    If fglbNew Then 'Ticket #29714 - Added the following back after the above ticket commented it out. Added back for new position esp. when it's multi position.
        SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_JOB='" & nJob & "' "
    Else
        SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_JOB='" & nJob & "' AND JH_SDATE=" & Date_SQL(dlpPosStDate)
    End If
    
    'If glbMultiGrid Then SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_GRID='" & clpGrid.Text & "' "
End If
dynaJobHIS.Filter = SQLQ
Screen.MousePointer = DEFAULT
If dynaJobHIS.BOF And dynaJobHIS.EOF Then
    glbStopSalary% = nCurrent
    Exit Function
Else
    glbStopSalary% = False
End If

If IsNull(dynaJobHIS("JH_WHRS")) Then fglbWhrs# = 0 Else fglbWhrs# = dynaJobHIS("JH_WHRS")
fglbJob$ = dynaJobHIS("JH_JOB") & ""     ' record
fglbSDate = dynaJobHIS("JH_SDATE") & ""
fglbGrid = dynaJobHIS("JH_GRID") & ""
fglbPayrollID = dynaJobHIS("JH_PAYROLL_ID") & ""
orgPosStDate = fglbSDate
If Not IsNull(dynaJobHIS("JH_JREASON")) Then
    fglbReason$ = dynaJobHIS("JH_JREASON")
End If
If Len(dynaJobHIS("JH_ID")) > 0 Then fglbJobID& = dynaJobHIS("JH_ID") Else fglbJobID& = 0

If Len(fglbGrid) > 0 And glbLambton Then txtLambtonJob = Left(fglbGrid, 1) & fglbJob$ & Mid(fglbGrid, 2)
'Hemu
If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls
    fglbNiagPhrs = clpCode(4).Text
    If fglbNiagPhrs = "" Then fglbNiagPhrs = 1
    fglbPhrs = dynaJobHIS("JH_PHRS")
Else
    fglbPhrs = dynaJobHIS("JH_PHRS")
End If

'City of Niagara Falls - Pick the Hours per Day from HRJOB table
If glbCompSerial = "S/N - 2276W" Then
    rsHRJOB.Open "SELECT JB_CODE, JB_DHRS FROM HRJOB WHERE JB_CODE = '" & dynaJobHIS("JH_JOB") & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRJOB.EOF Then
        If IsNull(rsHRJOB("JB_DHRS")) Or rsHRJOB("JB_DHRS") = "" Then
            fglbDhrs = dynaJobHIS("JH_DHRS")
        Else
            fglbDhrs = rsHRJOB("JB_DHRS")
        End If
    Else
        fglbDhrs = dynaJobHIS("JH_DHRS")
    End If
    rsHRJOB.Close
Else
    fglbDhrs = dynaJobHIS("JH_DHRS")
End If
'Hemu

'Ticket #24482 - Town of Caledon - Using hte VGroup field to store the Job's Division to create uniqueness between
'multiple same Position and Start Date positions linked to Salary.
If glbCompSerial = "S/N - 2182W" Then
    fglbDiv = dynaJobHIS("JH_DIV")
End If

dynaJobHIS.Filter = ""
Set_Position = True
Screen.MousePointer = DEFAULT

Exit Function

SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_HISTORY", "SELECT")
Resume Next

Exit Function

End Function

Private Sub Set_Current_Flag()
Dim SQLQ As String, Msg$
Dim dyn_HRSALHIS As New ADODB.Recordset

On Error GoTo SCFError

If glbMulti Then Exit Sub

'Hemu - 07/07/2003 Begin - Commented out the clone line cause it was giving Error
'                          as 'Row cannot be located for updating'
'Set dyn_HRSALHIS = Data1.Recordset.Clone
dyn_HRSALHIS.Open Data1.RecordSource, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'Hemu- 07/07/2003  End

Screen.MousePointer = HOURGLASS

If dyn_HRSALHIS.RecordCount < 1 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

If dyn_HRSALHIS.RecordCount > 0 Then dyn_HRSALHIS.MoveFirst
dyn_HRSALHIS("SH_CURRENT") = True
dyn_HRSALHIS.Update

Do Until dyn_HRSALHIS.EOF
    dyn_HRSALHIS.MoveNext
    If dyn_HRSALHIS.EOF Then Exit Do
    
    'Hemu - 07/07/2003 Begin - to improve speed, Jaddy suggested
    If dyn_HRSALHIS("SH_CURRENT") <> 0 Then
        dyn_HRSALHIS("SH_CURRENT") = False
        dyn_HRSALHIS.Update
    End If
    'Hemu - 07/07/2003 End
Loop
dyn_HRSALHIS.Close

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

Screen.MousePointer = DEFAULT

Exit Sub

SCFError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_SALARY_HISTORY", "Add")
Resume Next

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

chkCurrent.Enabled = TF
cmdChPos.Enabled = TF
comPayPer.Enabled = TF
comSalScale.Enabled = TF
fraSalary.Enabled = TF
medAmtChng(1).Enabled = TF
medAmtChng(2).Enabled = TF
medAmtChng(3).Enabled = TF
medPercentChng(1).Enabled = TF
medPercentChng(2).Enabled = TF
medPercentChng(3).Enabled = TF
 clpPostCode.Enabled = TF
dlpPosStDate.Enabled = TF
If glbCompSerial = "S/N - 2259W" Then 'Ticket #16877
    'force user to use the step to bring Salary
    If Not glbSetSal Then 'Ticket #17400 Use can edit it on Previous Salary screen
        medsalary.Enabled = False
    End If
Else
    medsalary.Enabled = TF
End If
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
dlpDate(0).Enabled = TF
clpCode(6).Enabled = TF 'WFC Currency Ticket #29244 Franks 09/22/2016

'Release 8.1 - County of Wellington - Grey out Next Review Date
If glbCompSerial = "S/N - 2262W" Then
    dlpDate(1).Enabled = False
Else
    dlpDate(1).Enabled = TF
End If
txtComment.Enabled = TF
cmbMarketLine.Enabled = TF
optUserSys(0).Enabled = TF
optUserSys(1).Enabled = TF
mskCampa.Enabled = TF
chkRedCircled.Enabled = TF 'Ticket #20648
If glbSetSal Or glbMulti Then
    clpPostCode.Enabled = TF
    If glbMulti Then
        dlpPosStDate.Enabled = TF
        clpGrid.Enabled = TF
    End If
    cmdChPos.Visible = False
Else
    clpPostCode.Enabled = False
    dlpPosStDate.Enabled = False
    clpGrid.Enabled = False
End If
' danielk - 01/06/2003 - added function to enable editing SH_WHRS for historical records (Ticket #3405)
' danielk - 01/07/2003 - don't enable, only disable in this function, enabling happen w/ edit pos/date btn
If TF = False Then txtWHRS.Enabled = False
' danielk - 01/06/2003 - end
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdRecal.Enabled = False
    cmdChPos.Enabled = False
End If
If glbtermopen Then
    cmdRecal.Enabled = False
'    cmdOK.Enabled = False
'    cmdCancel.Enabled = False
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdChPos.Visible = False
End If
'If Not gSec_Inq_Performance Then cmdPerform.Enabled = False
'If Not gSec_Inq_Position Then cmdPosition.Enabled = False
If glbLinamar Then
    Dim rsTB As New ADODB.Recordset
    rsTB.Open "SELECT ED_EMP FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        If rsTB!ED_EMP = "TEMP" Then
'            cmdNew.Enabled = False
'            cmdModify.Enabled = False
'            cmdDelete.Enabled = False
        End If
    End If
    rsTB.Close
End If

End Sub

Private Sub medTotal_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTotal_LostFocus()
    Call setPayPeriodSalary
End Sub

Sub MskCampa_GotFocus() 'Jaddy 8/9/99
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub mskCampa_LostFocus()
    Call Set_WFC_COMPA
End Sub
Private Sub OptUserSys_Click(Index As Integer) 'Jaddy 8/9/99
End Sub

Private Sub optUserSys_LostFocus(Index As Integer)
    txtUserSys = IIf(optUserSys(0), "", "U")
End Sub

Private Sub optUserSys_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 mskCampa.Visible = optUserSys(1)
End Sub


Private Sub scrControl_Change()
    panDetails.Top = 0 - scrControl.Value
End Sub

Private Sub txtComment_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmbMarketLine_GotFocus()   'Jaddy 8/9/99
    Call SetPanHelp(ActiveControl)
End Sub
Private Sub cmbMarketLine_LostFocus()
    txtMarketLine = cmbMarketLine
End Sub

Private Sub txtFiscalYear_LostFocus()
If Len((txtFiscalYear)) > 0 Then
    If Not IsNumeric(txtFiscalYear) Then
        MsgBox "Invalid Fiscal Year."
        txtFiscalYear.SetFocus
    End If
    If Val(txtFiscalYear) < 1900 Or Val(txtFiscalYear) > 3000 Then
        MsgBox "Invalid Fiscal Year."
        txtFiscalYear.SetFocus
    End If
End If
Call Set_MarketLine_List
Call Set_SalState
End Sub

Private Sub txtMarketLine_Change() 'Jaddy 8/9/99
  'cmbMarketLine.Clear
  'MarketLine_AddItem Me
  'setMarketLine Me
  Call SalMarketLineDesc
  Call Set_SalState
End Sub

Private Sub txtPosCode_LostFocus()

End Sub

Private Sub Set_COMPA()
Dim SQLQ As String, Msg As String
Dim iRec As Integer
Dim ssalary As Double
Dim X!, cX$
Dim ESalaryCode$
Dim HoursPerWeek!
Dim Compa!
Dim z%
Dim xsSalary  As Double
On Error GoTo UpRel_Err

If glbWFC And UnionExecNone Then Exit Sub

ESalaryCode$ = lblSalCode
'added by Bryan 22/sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
        medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
    ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
        medTotal.Text = medsalary.Text
    End If
    If Len(medTotal) = 0 Then
        ssalary = 0
    Else
        ssalary = medTotal
    End If
Else
    If Len(medsalary) = 0 Then
        ssalary = 0
    Else
        ssalary = medsalary
    End If
End If
HoursPerWeek! = Val(lblWhrs)

If ESalaryCode$ = "H" Then
    If ssalary > 500 Then
        MsgBox "Check if salary is paid Hourly or Annually"
        Exit Sub
    End If
End If
 
z% = getJOB(clpPostCode.Text, clpGrid.Text)
lblBANDCode = fglbBAND
Compa! = 0
If JobSnaps_PayScale(JobSnap_MidPoint!) <> 0 Then

    If JobSnaps_Salary_Code$ = "H" Then
        If ESalaryCode$ = "H" Then
            xsSalary = ssalary
        ElseIf ESalaryCode$ = "M" Then
            If HoursPerWeek! = 0 Then
                xsSalary = 0
            Else
                xsSalary = ((ssalary * 12) / HoursPerWeek!) / 52
            End If
        ElseIf ESalaryCode$ = "A" Then
            If HoursPerWeek! = 0 Then
                xsSalary = 0
            Else
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    xsSalary = (ssalary)
                Else
                xsSalary = (ssalary / HoursPerWeek!) / 52
                End If
            End If
        End If
    ElseIf JobSnaps_Salary_Code$ = "A" Then
        If ESalaryCode$ = "H" Then
            If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                xsSalary = (ssalary)
            Else
            xsSalary = (ssalary * HoursPerWeek!) * 52
            End If
        ElseIf ESalaryCode$ = "M" Then
            xsSalary = ssalary * 12
        ElseIf ESalaryCode$ = "A" Then
            xsSalary = ssalary
        End If
    End If
    Compa! = (xsSalary / JobSnaps_PayScale(JobSnap_MidPoint!)) * 100
End If
If glbFrench Then
    If Compa! > "999,99" Then Compa! = "999,99"
Else
    If Compa! > 999.99 Then Compa! = 999.99
End If
If glbCompSerial = "S/N - 2291W" Then Compa! = Round(Compa!, 0) 'Syndesis

lblCompaNum = Compa!

Exit Sub

UpRel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SAL HISTORY", "HR_SALARY_HISTORY", "INSERT")
Resume Next

End Sub

Private Sub Upd_Salary()    'RAUBREY 6/6/97
Dim X%
'Hemu - essex
'medSalary = Round2DEC(Val(orgSalary) + CCur(Val(medAmtChng(1))) + CCur(Val(medAmtChng(2))) + CCur(Val(medAmtChng(3))))
If glbFrench Then
    medsalary = Round2DEC(orgSalary1 + IIf(fglbAmtOld(1) <> CCur(medAmtChng(1)), CCur(medAmtChng(1)) - fglbAmtOld(1), 0) + IIf(fglbAmtOld(2) <> CCur(medAmtChng(2)), CCur(medAmtChng(2)) - fglbAmtOld(2), 0) + IIf(fglbAmtOld(3) <> CCur(medAmtChng(3)), CCur(medAmtChng(3)) - fglbAmtOld(3), 0))
    If fglbAmtOld(1) <> CCur(medAmtChng(1)) Then
        fglbAmtOld(1) = CCur(medAmtChng(1))
    End If
    If fglbAmtOld(2) <> CCur(medAmtChng(2)) Then
        fglbAmtOld(2) = CCur(medAmtChng(2))
    End If
    If fglbAmtOld(3) <> CCur(medAmtChng(3)) Then
        fglbAmtOld(3) = CCur(medAmtChng(3))
    End If
    If IsNumeric(medsalary.Text) Then
        orgSalary1 = medsalary
    End If
    If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #22952 Franks 12/10/2012
        'do not use the following calculation
    Else
        If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
            medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
        ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
            medTotal.Text = medsalary.Text
        End If
    End If
Else
    medsalary = Round2DEC(Val(orgSalary1) + IIf(fglbAmtOld(1) <> CCur(Val(medAmtChng(1))), CCur(Val(medAmtChng(1))) - fglbAmtOld(1), 0) + IIf(fglbAmtOld(2) <> CCur(Val(medAmtChng(2))), CCur(Val(medAmtChng(2))) - fglbAmtOld(2), 0) + IIf(fglbAmtOld(3) <> CCur(Val(medAmtChng(3))), CCur(Val(medAmtChng(3))) - fglbAmtOld(3), 0))
    If fglbAmtOld(1) <> CCur(Val(medAmtChng(1))) Then
        fglbAmtOld(1) = CCur(Val(medAmtChng(1)))
    End If
    If fglbAmtOld(2) <> CCur(Val(medAmtChng(2))) Then
        fglbAmtOld(2) = CCur(Val(medAmtChng(2)))
    End If
    If fglbAmtOld(3) <> CCur(Val(medAmtChng(3))) Then
        fglbAmtOld(3) = CCur(Val(medAmtChng(3)))
    End If
    orgSalary1 = medsalary
    If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #22952 Franks 12/10/2012
        'do not use the following calculation
    Else
        If IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) Then
            medTotal.Text = CDbl(medsalary.Text) + CDbl(medPremium.Text)
        ElseIf IsNumeric(medsalary.Text) And IsNumeric(medPremium.Text) = False Then
            medTotal.Text = medsalary.Text
        End If
    End If
End If
'Hemu - essex
' -----

End Sub

Private Function updFollow(xType)   'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim Edit1 As Integer
Dim rsTT As New ADODB.Recordset

'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

'Ticket #21365
'If fglHredsem <> "" Then    'DATE Renewal IS NOW MANDATORY
If IsDate(fglHredsem) Then     'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'SREV'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(fglHredsem)
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
    If fglbNew And dlpDate(1).Text <> "" Then
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'SREV'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpDate(1).Text)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
        ' Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = "SREV"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
                
            rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='SREV'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                rsTT("TB_KEY") = "SREV"
                rsTT("TB_DESC") = "Salary Review"
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            Call Grant_FollowUpCode_Security(glbUserID, "SREV", "Salary Review")
        End If
        rsFollow.Close
        rsTB.Close
                
        updFollow = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And dlpDate(1).Text <> "" Then
        ' 5/2/2001 Add by Frank for no duplicated record of HR_FOLLOW_UP Begin
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'SREV' "
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpDate(1).Text)
        

        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
        ' Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = "SREV"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        
            rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='SREV'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                rsTT("TB_KEY") = "SREV"
                rsTT("TB_DESC") = "Salary Review"
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            Call Grant_FollowUpCode_Security(glbUserID, "SREV", "Salary Review")
        
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
  
    If fglbNew = False And Edit1 = True And dlpDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = dlpDate(1).Text
            dynHRAT("EF_FREAS") = "SREV"
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If fglHredsem <> dlpDate(1).Text Then
            'Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpDate(1).Text = "" Then
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
       ' Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpDate(1).Text = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function
'Private Sub txtPosStDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtPayrollID_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPayrollID_KeyPress(KeyAscii As Integer)
If glbVadim Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUserSys_Change()
optUserSys(1) = IIf(txtUserSys = "U", True, False)
optUserSys(0) = Not optUserSys(1)
End Sub

Private Sub txtVGroup_Change()
If glbCompSerial = "S/N - 2429W" Then 'North Perth Ticket #19209 Franks 05/18/2011
    cboVGRoup.ListIndex = getVGrpcno(txtVGroup.Text)
End If
If glbCompSerial = "S/N - 2437W" Then 'KN&V Ticket #21097 Franks 11/02/2011
    cboVGRoup.ListIndex = getVGrpcno(txtVGroup.Text)
End If
End Sub


Private Sub txtWHRS_Change()
    lblWhrs.Caption = txtWHRS.Text
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 And Not glbWFC Then
        'dlpDate(2).Text = Updstats(0)
    End If
    If Index = 2 Then
        lblUserDesc = GetUserDesc(Updstats(2))
    End If
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

'If KeyAscii = 9 Then ' if the tab key was struck
'    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdClose.SetFocus
'    End If
'End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X As Integer, SQLQ

Call Display_Value

fglbJob$ = clpPostCode.Text
Call getJOB(clpPostCode.Text, clpGrid.Text)
optUserSys(1) = IIf(txtUserSys = "U", True, False)
optUserSys(0) = Not optUserSys(1)
mskCampa.Visible = optUserSys(1) And optUserSys(1).Visible

'WDGPHU - Ticket #17324
If glbCompSerial = "S/N - 2411W" Then
    txtPosGroup.Text = GetJobData(clpPostCode.Text, "JB_GRPCD", "")
End If

'Ticket #20652 - Town of Aurora
If glbCompSerial = "S/N - 2378W" Then
    lblPosGrp.Visible = True
    lblPosGrp.Caption = GetTABLDesc("JBGC", GetJobData(clpPostCode.Text, "JB_GRPCD", ""))
Else
    lblPosGrp.Caption = ""
    lblPosGrp.Visible = False
End If

'Ticket #22058 - The older position code sometimees does not show the Position Description, this method will
'refresh it.
clpPostCode.RefreshDescription

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

Private Sub DecSetup()
If glbCompDecHR = 3 Then
    medsalary.Format = "#,##0.000;(#,##0.000)"
    medTotal.Format = "#,##0.000;(#,##0.000)"
    medPremium.Format = "#,##0.000;(#,##0.000)"
    medAmtChng(1).Format = "#,##0.000;(#,##0.000)"
    medAmtChng(2).Format = "#,##0.000;(#,##0.000)"
    medAmtChng(3).Format = "#,##0.000;(#,##0.000)"
    vbxTrueGrid.Columns(1).NumberFormat = "#,##0.000;(#,##0.000)"
End If
If glbCompDecHR = 4 Then
    medsalary.Format = "#,##0.0000;(#,##0.0000)"
    medTotal.Format = "#,##0.0000;(#,##0.0000)"
    medPremium.Format = "#,##0.0000;(#,##0.0000)"
    medAmtChng(1).Format = "#,##0.0000;(#,##0.0000)"
    medAmtChng(2).Format = "#,##0.0000;(#,##0.0000)"
    medAmtChng(3).Format = "#,##0.0000;(#,##0.0000)"
    vbxTrueGrid.Columns(1).NumberFormat = "#,##0.0000;(#,##0.0000)"
End If

'7.9 - Show Compa-Ration based on the Company Pref. setup
vbxTrueGrid.Columns(7).Visible = gsCompaRatio

End Sub

Private Sub setPercent()
Dim X%
Dim newAmtChg

If glbCompSerial = "S/N - 2436W" Then  'Family Day - Ticket #21152 Franks 04/01/2013
    Exit Sub
End If
If fglbEmptyNew Then
    medPercentChng(1) = 1
    medAmtChng(1) = medsalary
Else
    'Release 8.0 - Ticket #22682 - The % and Amount Change not getting computed correctly with fglbSHold@ as it is not
    'getting correct original value assigned. Also in the rest of the code the 'orgSalary' value is being used so not
    'sure why fglbSHold@ was being checked against to see if the salary has changed in the first place.
    'If fglbSHold@ <> CCur(medsalary) Then
    If CCur(orgSalary) <> CCur(medsalary) Then
        For X% = 2 To 3
            medPercentChng(X%) = 0
            medAmtChng(X%) = 0
        Next X%
        
        'Release 8.0 - Logic fix
        If orgSalCD1 = lblSalCode Then
            medAmtChng(1) = medsalary - orgSalary
            newAmtChg = medAmtChng(1)
        Else
            If orgSalCD1 = "H" And lblSalCode = "A" Then
                'Change the New Salary to Hourly
                'newAmtChg = Round2DEC((Val(medsalary) / 52) / Val(txtWHRS.Text))
                'medAmtChng(1) = newAmtChg - orgSalary
                
                'Change the Old Hourly Rate to Salary - Jerry want to show the amount change in current Per value
                newAmtChg = Round2DEC((Val(orgSalary) * Val(txtWHRS.Text)) * 52)
                medAmtChng(1) = medsalary - newAmtChg
                
                newAmtChg = medAmtChng(1)
            Else
                'Change the New Salary to Annual
                'newAmtChg = Round2DEC((Val(medsalary) * Val(txtWHRS.Text)) * 52)
                'medAmtChng(1) = newAmtChg - orgSalary
                
                'Change the Old Salary to Hourly - Jerry want to show the amount change in current Per value
                If IsNumeric(txtWHRS.Text) And Val(txtWHRS.Text) <> 0 Then  'Check for 0 value or blank value, otherwise gives Division by 0 error
                    newAmtChg = Round2DEC((Val(orgSalary) / 52) / Val(txtWHRS.Text))
                    medAmtChng(1) = medsalary - newAmtChg
                    
                    newAmtChg = medAmtChng(1)
                Else
                    medAmtChng(1) = 0
                    newAmtChg = 0
                End If
            End If
        End If
        'If medAmtChng(1) <> 0 Then
        '    If orgSalary <> 0 Then
        '        medPercentChng(1) = medAmtChng(1) / orgSalary
        '    Else
        '        medPercentChng(1) = 1
        '    End If
        'Else
        '    medPercentChng(1) = 0
        'End If
        If medAmtChng(1) <> 0 Then
            If orgSalary <> 0 Then
                'medPercentChng(1) = newAmtChg / orgSalary
                'Because of the logic change to show the amount change as current Per value
                If orgSalCD1 = "H" And lblSalCode = "A" Then
                    newAmtChg = Round2DEC((Val(medsalary) / 52) / Val(txtWHRS.Text))
                    newAmtChg = newAmtChg - orgSalary
                Else
                    'If lblSalCode = "A" Then
                    '    newAmtChg = Round2DEC((Val(medsalary) * Val(txtWHRS.Text)) * 52)
                    '    newAmtChg = newAmtChg - orgSalary
                    'End If
                End If
                'Converting OrgSalary to the current Per so the % of difference can be calculated
                If orgSalCD1 <> lblSalCode Then
                    If orgSalCD1 = "H" Then
                        'medPercentChng(1) = newAmtChg / (Round2DEC((Val(orgSalary) * Val(txtWHRS.Text)) * 52))
                        medPercentChng(1) = newAmtChg / orgSalary
                    ElseIf orgSalCD1 = "A" Then
                        medPercentChng(1) = newAmtChg / Round2DEC((Val(orgSalary) / 52) / Val(txtWHRS.Text))
                    End If
                Else
                    medPercentChng(1) = newAmtChg / orgSalary
                End If
            Else
                medPercentChng(1) = 1
            End If
        Else
            medPercentChng(1) = 0
        End If
    
    End If
End If
End Sub

Private Sub Get_OrgSalary()
Dim SQLQ As String, HRSH_Snap As New ADODB.Recordset
On Error GoTo JS_Err
SQLQ = "Select SH_SALARY from HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND SH_JOB = '" & clpPostCode.Text & "' "
'Hemu
SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(dlpPosStDate.Text)
'Hemu
SQLQ = SQLQ & " ORDER BY "

'Ticket #21511 - County of Oxford - since they are able to switch between multi and non-multi, they are
'seeing an issue with sort order, so this will fix it.
If glbCompSerial = "S/N - 2259W" Then
    SQLQ = SQLQ & " SH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
ElseIf glbMulti Then
    SQLQ = SQLQ & " SH_CURRENT " & IIf(glbSQL, "DESC", "") & ","
End If
SQLQ = SQLQ & " SH_EDATE DESC"

HRSH_Snap.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If HRSH_Snap.BOF And HRSH_Snap.EOF Then
    orgSalary = 0
    orgSalary1 = 0
Else
    orgSalary = HRSH_Snap("SH_SALARY")
    orgSalary1 = HRSH_Snap("SH_SALARY")
End If

HRSH_Snap.Close
Exit Sub

JS_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SALARY History Snap", "HR_SALARY_HISTORY", "SELECT")
Resume Next

End Sub

Sub DoWFCGrids(NewEmp As Boolean)
    Dim I As Integer
    
    ' dkostka - 08/31/2000 - WFC requested changes.
    If glbWFC Then
        Data3.ConnectionString = Data1.ConnectionString
        If glbtermopen Then
            Data3.RecordSource = "SELECT ED_ORG,ED_DIV FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            Data3.RecordSource = "SELECT ED_ORG,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
        End If
        Data3.Refresh
        
        If Format(Data3.Recordset("ED_ORG"), "@") = "EXEC" Or Format(Data3.Recordset("ED_ORG"), "@") = "NONE" Then
            UnionExecNone = True
            'If Not NewEmp And Data1.Recordset.EOF = False Then txtMarketLine.DataField = "SH_MARKETLINE"
            'MarketLine_AddItem Me
            
            If NewEmp Then
                If Len(txtMarketLine) = 0 Then 'Ticket# 8046
                    txtMarketLine = GetMarketlineFromDiv(Data3.Recordset("ED_DIV"))
                End If
            End If
            'lblBand.Top = 4500 '4350
            'lblBANDCode.Top = 4500 '4350
            lblBand.Left = 5370 + 1500
            lblBANDCode.Left = 6600 + 900
            lblBand.Visible = True
            lblBANDCode.Visible = True
            lblMarketLine.Visible = True
            cmbMarketLine.Visible = True
            lblMLine.Visible = True
            lblsalstate(0).Visible = True
            lblsalstate(1).Visible = True
            lblsalstate(2).Visible = True
            optUserSys(0).Visible = False 'True Ticket# 6962 WFC doesn't need it
            optUserSys(1).Visible = False 'True Ticket# 6962 WFC doesn't need it
            comSalScale.Visible = False
            lblTitle(9).Visible = False
            lblTitle(13).Visible = True
            lblSalaryGrade.Visible = False
            mskCampa.Visible = False 'True Ticket# 6962 WFC doesn't need it
            lblFiscalYear.Left = 5280 '7200
            txtFiscalYear.Left = 6300
            lblFiscalYear.Visible = True
            txtFiscalYear.Visible = True
            lblTitle(17).Visible = True
            dlpDate(2).Visible = True
            cmdTranDate.Visible = True
            lblPlant.Visible = True
            clpCode(0).Visible = True
            lblPlant.Left = 5280
            clpCode(0).Left = 5980
            fraSalary.Width = 6645 '5150
        Else
            UnionExecNone = False
            txtMarketLine.DataField = ""
            comSalScale.Clear
            'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
            'For I = 1 To 11
            'For I = 1 To 15
            For I = 1 To 20
                comSalScale.AddItem Format(I, "00")
            Next
            
            lblBand.Visible = False
            lblMarketLine.Visible = False
            cmbMarketLine.Visible = False
            lblMLine.Visible = False
            lblsalstate(0).Visible = False
            lblsalstate(1).Visible = False
            lblsalstate(2).Visible = False
            optUserSys(0).Visible = False
            optUserSys(1).Visible = False
            comSalScale.Visible = True
            lblTitle(9).Visible = True
            lblTitle(13).Visible = False
            lblSalaryGrade.Visible = True
            mskCampa.Visible = False
            lblFiscalYear.Visible = False
            txtFiscalYear.Visible = False
            lblTitle(17).Visible = False
            dlpDate(2).Visible = False
            cmdTranDate.Visible = False
            lblPlant.Visible = False
            clpCode(0).Visible = False
            fraSalary.Width = 9045
        End If
    Else
        ' Not WFC.
        txtMarketLine.DataField = ""
        comSalScale.Clear
        'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
        'For I = 1 To 11
        'For I = 1 To 15
        For I = 1 To 20
            If glbCompSerial = "S/N - 2366W" Then   'Family Youth Child Services of Muskoka
                If I = 1 Then
                    comSalScale.AddItem "Start"
                Else
                    comSalScale.AddItem Format(I - 1, "00")
                End If
            Else
                comSalScale.AddItem Format(I, "00")
            End If
        Next
        
        If Not glbSyndesis Then
            lblBand.Visible = False
            lblBANDCode.Visible = False
        End If
        
        lblMarketLine.Visible = False
        cmbMarketLine.Visible = False
        lblMLine.Visible = False
        lblsalstate(0).Visible = False
        lblsalstate(1).Visible = False
        lblsalstate(2).Visible = False
        optUserSys(0).Visible = False
        optUserSys(1).Visible = False
        If (glbCompSerial = "S/N - 2351W") Then 'Burlington Tech
            comSalScale.Visible = False
            lblTitle(9).Visible = False
        Else
            comSalScale.Visible = True
            lblTitle(9).Visible = True
        End If
'        comSalScale.Visible = True
'        lblTitle(9).Visible = True
        lblTitle(13).Visible = False
        lblSalaryGrade.Visible = False
        mskCampa.Visible = False
    End If
    
    ' dkostka end
End Sub

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ


'Hemu - 10/09/2003 Begin
'Ticket #21511
If Not Data1.Recordset.EOF Then
    If glbSetSal Then
        Call CR_JobHis_Snap(False)
    Else
        Call CR_JobHis_Snap(IIf(Data1.Recordset!SH_CURRENT, True, False))
    End If
Else
    Call CR_JobHis_Snap(False)
End If

'Call Set_Position(fglbJob$, False)

If Not Data1.Recordset.EOF Then 'Ticket #13062 Frank 05/10/2007
    If Not IsNull(Data1.Recordset("SH_JOB")) Then
        Call Set_Position(Data1.Recordset("SH_JOB"), False)
        
        If glbCompSerial = "S/N - 2172W" Then 'Lanark Ticket #17221
            Call Set_SalaryLevel(Data1.Recordset("SH_JOB"))
        End If
    End If
End If
clpPostCode.seleEMPCode = fglbJobList
'Hemu - 10/09/2003 End

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    
    lblPayPeriodSalary = ""
    lblHoursPay = ""
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    If glbtermopen Then
        If glbCompSerial = "S/N - 2191W" Then
            SQLQ = SQLQ & " SELECT *,IIF(ISNULL(JB_DESCR2),SH_GRADE,IIF(JB_DESCR2<>'.5' OR SH_GRADE='00', VAL(SH_GRADE),(VAL(SH_GRADE)+1)/2)) AS SH_GRADESHOW "
            SQLQ = SQLQ & " FROM Term_SALARY_HISTORY "
            SQLQ = SQLQ & " LEFT JOIN HRJOB ON Term_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
            SQLQ = SQLQ & " WHERE SH_ID = " & Data1.Recordset!sh_id
            'vbxTrueGrid.Columns(5).NumberFormat = "0.0"
            vbxTrueGrid.Columns(6).NumberFormat = "0.0"
        ElseIf glbOracle Then
            SQLQ = SQLQ & "SELECT Term_SALARY_HISTORY.*,SH_GRADE AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
        Else
            SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW FROM Term_SALARY_HISTORY "
        End If
        SQLQ = SQLQ & "WHERE SH_ID = " & Data1.Recordset!sh_id
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        If glbCompSerial = "S/N - 2191W" Then
            SQLQ = SQLQ & " SELECT *,IIF(ISNULL(JB_DESCR2),SH_GRADE,IIF(JB_DESCR2<>'.5' OR SH_GRADE='00', VAL(SH_GRADE),(VAL(SH_GRADE)+1)/2)) AS SH_GRADESHOW "
            SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
            SQLQ = SQLQ & " LEFT JOIN HRJOB ON HR_SALARY_HISTORY.SH_JOB=HRJOB.JB_CODE "
            'vbxTrueGrid.Columns(5).NumberFormat = "0.0"
            vbxTrueGrid.Columns(6).NumberFormat = "0.0"
        ElseIf glbOracle Then
            SQLQ = SQLQ & "SELECT HR_SALARY_HISTORY.*, SH_GRADE AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
        Else
            SQLQ = SQLQ & "SELECT *,LTRIM(SH_GRADE) AS SH_GRADESHOW FROM HR_SALARY_HISTORY "
            
        End If
        SQLQ = SQLQ & " WHERE SH_ID = " & Data1.Recordset!sh_id
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If

    If rsDATA.EOF Or rsDATA.BOF Then
        'Hemu - The buttons on the toolbar was not enabling properly if multiple forms
        'were open
        If flgloaded Then
            If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                'Hemu
                Call SET_UP_MODE
            End If
        End If
        
        Exit Sub
    End If
    
    Call Set_Control("R", Me, rsDATA)
    
    'Hemu - 08/11/2003 Begin - Calculate and Display Salary per Pay Period
    If Not IsNull(Data1.Recordset("SH_JOB")) Then   'Ticket #26419 - had to add this from above here again, it was not picking the right Pay Period value
        Call Set_Position(Data1.Recordset("SH_JOB"), False) 'Ticket #26419 - had to add this from above here again, it was not picking the right Pay Period value
    End If
    Call setPayPeriodSalary
    'Hemu - 08/11/2003 End
    
    If glbCompSerial = "S/N - 2359W" Then
        clpCode(5) = txtComment
    End If
End If
    
If glbLambton Then
    If Len(clpGrid.Text) > 0 And Len(clpPostCode.Text) Then
        txtLambtonJob = Left(clpGrid, 1) & clpPostCode & Mid(clpGrid, 2)
    End If
End If

If glbCompSerial = "S/N - 2373W" Then
    If txtVGroup <> "" Then
        cboVGRoup = txtVGroup
    Else
        cboVGRoup = ""
    End If
    If txtVStep <> "" Then
        cboVStep = txtVStep
    Else
        cboVStep = ""
    End If
End If

If glbWFC Then  'Ticket #13581
    If locCountry = "AUSTRALIA" Then
        If Trim(comPayPer.Text) = "Hour" Then
            medsalary.Format = "#,##0.0000;(#,##0.0000)"
            locCompDecHR = 4
        Else
            medsalary.Format = "#,##0.00;(#,##0.00)"
            locCompDecHR = 2
        End If
    Else
        medsalary.Format = "#,##0.00;(#,##0.00)"
        locCompDecHR = glbCompDecHR
    End If
    
    'see WFC Pension Outstanding Tasks By Nov0609.doc in W:\2008 Projects\Pension\Pension Phase II
    'November 05, 2009 - Support/Non Billable Items
    If lblSalCode.Caption = "A" Or lblSalCode.Caption = "M" Then
        lblTitle(10).FontBold = True
    Else
        lblTitle(10).FontBold = False
    End If
End If

Call SET_UP_MODE

Me.cmdModify_Click
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
UpdateRight = gSec_Upd_Salary
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
    cmdRecal.Enabled = False
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    cmdRecal.Enabled = False
    TF = False
Else
    UpdateState = OPENING
    TF = True
    cmdRecal.Enabled = True
End If

Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
If Not Updateble Then TF = False
Call ST_UPD_MODE(TF)
End Sub


Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmESALARY.Caption = "Salary - " & Left$(glbLEE_SName, 5)
    frmESALARY.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub setGridList(nJob)
    
Dim rsGrid As New ADODB.Recordset
Dim xGridList As String
Dim SaveGrid As String
If Not glbMultiGrid Then Exit Sub
SaveGrid = clpGrid
clpGrid = ""
If Len(clpPostCode.Text) > 0 Then
    rsGrid.Open "SELECT JB_ID,JB_GRID FROM HRJOB_GRADE WHERE JB_CODE='" & CStr(nJob) & "'", gdbAdoIhr001, adOpenForwardOnly
    xGridList = ""
    Do Until rsGrid.EOF
        xGridList = xGridList & "," & rsGrid("JB_GRID")
        rsGrid.MoveNext
    Loop
    If xGridList <> "" Then xGridList = Mid(xGridList, 2)
    clpGrid.seleEMPCode = xGridList
    rsGrid.Close
Else
    clpGrid.seleEMPCode = "NONE-GRID"
End If
clpGrid = SaveGrid
End Sub

Private Sub setDivList(nJob, nStartDate)
    Dim rsDiv As New ADODB.Recordset
    Dim xDivList As String
    Dim SaveDiv As String
    
    If Not glbMulti Then Exit Sub
    
    SaveDiv = clpDiv
    clpDiv = ""
    If Len(clpPostCode.Text) > 0 Then
        rsDiv.Open "SELECT JH_ID,JH_DIV FROM HR_JOB_HISTORY WHERE JH_JOB='" & CStr(nJob) & "' AND JH_SDATE = " & Date_SQL(nStartDate), gdbAdoIhr001, adOpenForwardOnly
        xDivList = ""
        Do Until rsDiv.EOF
            xDivList = xDivList & "," & rsDiv("JH_DIV")
            rsDiv.MoveNext
        Loop
        If xDivList <> "" Then xDivList = Mid(xDivList, 2)
        clpDiv.seleEMPCode = xDivList
        rsDiv.Close
    Else
        clpDiv.seleEMPCode = "NONE-DIVISION"
    End If
    clpDiv = SaveDiv
End Sub

Private Sub UpdatePTAdministeredBy(mPT, mAdministeredBy) 'for CCAC London saving Client transfer pop-up window's info
    gdbAdoIhr001.Execute "update HREMP set ED_PT='" & mPT & "', ED_ADMINBY='" & mAdministeredBy & "' where ED_EMPNBR=" & lblEENum
End Sub

Private Function GetMarketlineFromDiv(xDiv)
Dim rsODiv As New ADODB.Recordset
Dim SQLQ, xStr
    xStr = ""
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
    rsODiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsODiv.EOF Then
        If Not IsNull(rsODiv("DV_MARKETLINE")) Then
            xStr = rsODiv("DV_MARKETLINE")
        End If
    End If
    rsODiv.Close
    GetMarketlineFromDiv = xStr
End Function

Private Function VGroupList() As String
Dim retVal As String, ctyFile
retVal = ""
ctyFile = glbIHRREPORTS & "VGroupList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, retVal
    Close #1
End If

ResumeHere:
If InStr(retVal, cboVGRoup) = 0 And cboVGRoup <> "" Then
    retVal = retVal & "&" & cboVGRoup
    cboVGRoup.AddItem cboVGRoup
End If
Open ctyFile For Output As #1
Print #1, retVal
Close #1
VGroupList = retVal
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt VGroupList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Function VStepList() As String
Dim retVal As String, ctyFile
retVal = ""
ctyFile = glbIHRREPORTS & "VStepList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, retVal
    Close #1
End If

ResumeHere:
If InStr(retVal, cboVStep) = 0 And cboVStep <> "" Then
    retVal = retVal & "&" & cboVStep
    cboVStep.AddItem cboVStep
End If
Open ctyFile For Output As #1
Print #1, retVal
Close #1
VStepList = retVal
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt VStepList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function


Private Sub ResetFlagAudit()
On Error GoTo Eh
Dim strSQL As String
Dim rs As New ADODB.Recordset

    strSQL = "SELECT AU_ID,AU_UPLOAD FROM HRAUDIT WHERE AU_EmpNBR=" & glbLEE_ID
    strSQL = strSQL & " AND AU_LDATE=" & Date_SQL(dlpDate(0).Text)
    strSQL = strSQL & " AND AU_SREASON = '" & clpCode(1).Text & "'"
    strSQL = strSQL & " AND AU_SALARY = " & medsalary
    'rs.Open strSQL, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic, adCmdText
    rs.Open strSQL, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
    If rs.EOF = False And rs.BOF = False Then
        rs.MoveFirst
        Do While Not rs.EOF
            rs("AU_UPLOAD") = "Y"
            rs.Update
            rs.MoveNext
        Loop
    End If
    rs.Close
    
exH:
    Set rs = Nothing
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Updating AUDIT RECORD", "AUDIT FILE", "UPDATE")
    Call RollBack '28July99 js
    Resume exH

    
End Sub
Private Sub ChangeEDateAudit(xEDate)
On Error GoTo Eh
Dim strSQL As String
Dim rs As New ADODB.Recordset

    strSQL = "SELECT AU_LDATE, AU_SEDATE FROM HRAUDIT WHERE AU_EMPNBR=" & glbLEE_ID
    strSQL = strSQL & " AND AU_LDATE=" & Date_SQL(xEDate)
    strSQL = strSQL & " AND AU_SREASON = '" & clpCode(1).Text & "'"
    strSQL = strSQL & " AND AU_SALARY = " & medsalary
    rs.Open strSQL, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic, adCmdText

    If rs.EOF = False And rs.BOF = False Then
        rs("AU_LDATE") = dlpDate(0).Text
        rs("AU_SEDATE") = dlpDate(0).Text
        rs.Update
    End If
    rs.Close
    
exH:
    Set rs = Nothing
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Updating AUDIT RECORD", "AUDIT FILE", "UPDATE")
    Call RollBack '28July99 js
    Resume exH

    
End Sub
Private Function fgetSection(xID)
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If glbtermopen Then
        strSQL = "SELECT ED_SECTION FROM TERM_HREMP WHERE TERM_SEQ =" & xID
        rs.Open strSQL, gdbAdoIhr001X, adOpenStatic, adLockOptimistic, adCmdText
    Else
        strSQL = "SELECT ED_SECTION FROM HREMP WHERE ED_EMPNBR =" & xID
        rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    End If
    
    If rs.EOF = False Then
        If Not IsNull(rs("ED_SECTION")) Then
            fSection = rs("ED_SECTION")
        Else
            fSection = ""
        End If
    End If
    rs.Close
    Set rs = Nothing
    

End Function

Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
    If gsEMAIL_ONSALARY Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONSALARY")
            
        'Ticket #18090 - begin
        If glbCompSerial = "S/N - 2382W" Then  'Samuel
            xToEmail = GetComPreferEmail("EMAIL_ONSALARY", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONSALARY")
            End If
        Else
            'Ticket #20317 - Send email to More Emails list as well.
            xToEmail = GetComPreferEmail("EMAIL_ONSALARY", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONSALARY")
            End If
        End If
        'Ticket #18090 - end
            
        'If Len(xEmail) > 0 Then    'Hemu - (Ticket #11562) - Jerry asked to remove the check for email address presence.
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            
            'Samuel Ticket #18352, do not cc it to employee
            'Ticket #18856 - Friesens Corporation - do not cc it to the employee
            If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2279W" Then
            Else
                frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            End If
            'frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice"
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            'End If
        '    MsgBox "There is no email address for the 'Email Notification on Salary' on Company Preference screen. "
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

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    If gsEMAIL_ONSALARY Then
        If Not UserEmailExist Then
            Exit Sub
        End If

        xToEmail = GetComPreferEmail("EMAIL_ONSALARY", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONSALARY")
        End If
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352, do not cc it to employee
            'Else
            '    frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            'End If
            'frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice"
            'Ticket #18578
            'frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice - " & lblEEName.Caption
            'Ticket #18755
            xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
            If Len(xBranch) > 0 Then
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR Salary Change Notice - " & xBranch & lblEEName.Caption
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

    End If
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


Private Function getLastPayP() As String
On Error GoTo Eh

Dim SQLQ As String
Dim rs As New ADODB.Recordset

SQLQ = "SELECT SH_PAYP FROM HR_SALARY_HISTORY WHERE SH_EMPNBR=" & glbLEE_ID & " AND SH_CURRENT<>0"
rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText

If rs.EOF = False And rs.BOF = False Then
    If Not IsNull(rs("SH_PAYP")) Then
        getLastPayP = rs("SH_PAYP")
    Else
        getLastPayP = ""
    End If
Else
    getLastPayP = ""
End If

exH:
    Exit Function
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get Last PayPeriod", "HR_SALARY_HISTORY", "Last Pay Period")
    Resume Next
End Function

Private Sub Delete_Existing_Vadim(xUpdType)
    'For Vadim only - Delete the existing entries Ticket #15542
    'If glbVadim And (CVDate(Format(dlpDate(0), "mm/dd/yyyy")) > CVDate(Format(Date, "mm/dd/yyyy"))) Then
        Dim xSalRate, xPayRate
        Dim Salary_PayCode
        Dim PayCodeInfo As PayCodeInfoType
    
        xSalRate = 0
        xPayRate = 0
        If glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
            If xUpdType = "D" Then
                Call Compute_Salary_Vadim_Based(glbLEE_ID, lblSalCode, medsalary, fglbNiagPhrs, fglbDhrs, xSalRate, xPayRate)
            ElseIf xUpdType = "C" Then
                Call Compute_Salary_Vadim_Based(glbLEE_ID, OSalCD, OSalary, fglbNiagPhrs, fglbDhrs, xSalRate, xPayRate)
            End If
        Else
            If xUpdType = "D" Then
                Call Compute_Salary_Vadim_Based(glbLEE_ID, lblSalCode, medsalary, fglbPhrs, fglbWhrs, xSalRate, xPayRate)
            ElseIf xUpdType = "C" Then
                Call Compute_Salary_Vadim_Based(glbLEE_ID, OSalCD, OSalary, fglbPhrs, fglbWhrs, xSalRate, xPayRate)
            End If
        End If
        
        'Call functions to delete the entries in the SY_INTERFACE & SY_INTERFACE_BATCH tables and HR_VADIM_SY_INTERFACE
        If xUpdType = "D" Then
            Call Update_HR_Vadim_Sy_Interface("", glbLEE_ID, OEDate, xPayRate, oJob, xUpdType, Format(Now, "mm/dd/yyyy h:m:s"), "Pay Rate Delete")
        ElseIf xUpdType = "C" Then
            Call Update_HR_Vadim_Sy_Interface("", glbLEE_ID, OEDate, xPayRate, oJob, xUpdType, Format(Now, "mm/dd/yyyy h:m:s"), "Pay Rate Change")
        End If
        
        Call getPayCodeInfo(PayCodeInfo, "SALARY")
        If Len(PayCodeInfo.PayCode) > 0 Then
            Salary_PayCode = PayCodeInfo.PayCode
            If isExistTransCode(txtPayrollID.Text, Salary_PayCode) = 1 Then
                If getPayType(glbLEE_ID) = "S" Then
                    If xUpdType = "D" Then
                        Call Update_HR_Vadim_Sy_Interface("", glbLEE_ID, OEDate, xSalRate, oJob, xUpdType, Format(Now, "mm/dd/yyyy h:m:s"), "Salary Delete")
                    ElseIf xUpdType = "C" Then
                        Call Update_HR_Vadim_Sy_Interface("", glbLEE_ID, OEDate, xSalRate, oJob, xUpdType, Format(Now, "mm/dd/yyyy h:m:s"), "Salary Delete")
                    End If
                End If
            End If
        End If
    'End If
End Sub
Private Function GetDivisionCode()
    Dim SQLQ As String, xDivCode
    Dim rsXEMP As New ADODB.Recordset
    GetDivisionCode = ""
    locCountry = ""
    SQLQ = "SELECT ED_DIV FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsXEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsXEMP.EOF Then
        xDivCode = rsXEMP("ED_DIV")
    End If
    rsXEMP.Close
    GetDivisionCode = xDivCode
End Function

Private Sub TabOrderSetup()
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18090
        clpPostCode.TabIndex = 1
        dlpPosStDate.TabIndex = 2
        txtPosGroup.TabIndex = 3
        clpCode(1).TabIndex = 4
        medsalary.TabIndex = 5
        comPayPer.TabIndex = 6
        dlpDate(0).TabIndex = 7
        clpCode(4).TabIndex = 8
    End If
End Sub

Private Function AUDIT_NGS_TRANS()
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
Dim xDate1, xDate2
Dim xlocFlag As Boolean
Dim xOldVal, xNewVal


On Error GoTo AUDIT_ERR
If Not glbNGS_OnFlag Then
    AUDIT_NGS_TRANS = True
    Exit Function
End If

SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsEmpee.EOF Then
    Exit Function
Else
    If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
    If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
    If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
End If
rsEmpee.Close

'No NGS Sub Group, skip
If Len(glbWFCNGSSubGroup) = 0 Then Exit Function

'Ticket #20385 Franks 05/31/2011
'Change IHR to write to the NGS Audit Table if the employee has a NGS Sub Group
'regardless of entering a Start Date.
''xNGSStart = ""
''SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
''rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
''If Not rsEmpOther.EOF Then
''    If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
''        xNGSStart = rsEmpOther("ER_OTHERDATE1")
''    End If
''End If
''rsEmpOther.Close
'''No NGS Effective Date, skip
''If Len(xNGSStart) = 0 Then Exit Function

'Ticket #20385 Franks 05/31/2011
xLDate = Date
''xLDate = dlpDate(0).Text ' Salary Effective Date
''If IsDate(xNGSStart) Then
''    If CVDate(xNGSStart) > CVDate(xLDate) Then
''        xLDate = CVDate(xNGSStart)
''    End If
''End If

'NGS field changes --------------------------------------
xlocFlag = False
If OEDate <> dlpDate(0).Text Then
    If Len(OEDate) = 0 Then
        xlocFlag = True
    Else
        If Not (CVDate(OEDate) = CVDate(dlpDate(0).Text)) Then
            xlocFlag = True
        End If
    End If
    If xlocFlag Then
        xDate1 = OEDate
        xDate2 = dlpDate(0).Text
        Call NGSAuditAdd(glbLEE_ID, "M", "Salary History", "Effective Date", xDate1, xDate2, xLDate)
    End If
End If

'Salary amount
xOldVal = OSalary
xNewVal = medsalary
If Not (OSalary = medsalary) Then
    If xOldVal = 0 Then xOldVal = ""
    Call NGSAuditAdd(glbLEE_ID, "M", "Salary History", "Salary", xOldVal, xNewVal, xLDate)
End If


AUDIT_NGS_TRANS = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING NGS AUDIT RECORD", "NGS AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me
    
End Function

Private Sub SetDefaultsSamuel()
Dim rsEmpp As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT ED_EMPNBR, ED_ADMINBY, ED_PT, ED_REGION, ED_SECTION FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    rsEmpp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpp.EOF Then
        If rsEmpp("ED_PT") = "FTC" Or rsEmpp("ED_PT") = "PTC" Then
            'comPayPer.ListIndex = 1 'Hourly
            lblSalCode.Caption = "H"
            clpCode(4).Text = "V" 'Pay Period
        Else
            If rsEmpp("ED_ADMINBY") = "5231" Or rsEmpp("ED_ADMINBY") = "5230" Or rsEmpp("ED_ADMINBY") = "5232" Then
                'comPayPer.ListIndex = 0 'Annum
                'lblSalCode.Caption = "A"
                'clpCode(4).Text = "S" 'Pay Period
                If rsEmpp("ED_PT") = "ST" Then 'Ticket #25696 Franks 07/10/2014 - student
                    lblSalCode.Caption = "H"
                    clpCode(4).Text = "V" 'Pay Period
                Else
                    lblSalCode.Caption = "A"
                    clpCode(4).Text = "S" 'Pay Period
                End If
            End If
            If rsEmpp("ED_ADMINBY") = "5322" Or rsEmpp("ED_ADMINBY") = "2158" Then
                'comPayPer.ListIndex = 1 'Hourly
                lblSalCode.Caption = "H"
                clpCode(4).Text = "H" 'Pay Period
                
                'Ticket #20695 for Samuel Franks 09/23/2011
                If Not IsNull(rsEmpp("ED_REGION")) And Not IsNull(rsEmpp("ED_SECTION")) Then
                    If rsEmpp("ED_ADMINBY") = "5322" And rsEmpp("ED_REGION") = "RFG" And rsEmpp("ED_SECTION") = "31" Then
                        clpCode(4).Text = "V" 'Pay Period
                    End If
                End If
            End If
        End If
    End If
    rsEmpp.Close
End Sub

Public Sub getNONsecuritiesAgain(xEmpNo)
Dim rsLEmp As New ADODB.Recordset
Dim SQLQ As String

    'get glbNoEXEC and glbNoNONE again
    Call Dept_Secure
    
    'get glbUNION again
    SQLQ = "SELECT ED_EMPNBR, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsLEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLEmp.EOF Then
        If Not IsNull(rsLEmp("ED_ORG")) Then
            glbUNION = rsLEmp("ED_ORG")
        Else
            glbUNION = ""
        End If
    End If
    rsLEmp.Close
    
End Sub

Private Sub ScreenDMuskoka()
Dim vList
Dim X
'Added by Bryan 23/Sep/05 Ticket#9343
cboVGRoup.Clear
cboVStep.Clear
vList = VGroupList
X = 1
Do While X > 0
    X = InStr(vList, "&")
    If X > 0 Then
        cboVGRoup.AddItem Left(vList, X - 1)
        vList = Mid(vList, X + 1)
    Else
        cboVGRoup.AddItem vList
    End If
Loop
vList = VStepList
X = 1
Do While X > 0
    X = InStr(vList, "&")
    If X > 0 Then
        cboVStep.AddItem Left(vList, X - 1)
        vList = Mid(vList, X + 1)
    Else
        cboVStep.AddItem vList
    End If
Loop
End Sub
Private Sub ScreenNorthPerth() 'North Perth Ticket #19209 Franks 05/18/2011
Dim I As Integer
    lblTitle(16).Visible = False
    medPremium.Visible = False
    lblTitle(19).Caption = "Pay Type"
    txtVGroup.DataField = "SH_VGROUP"
    cboVGRoup.Clear
    'Ticket #21232 Franks 01/26/2012 - begin
    'cboVGRoup.AddItem = "HOURLY"
    'cboVGRoup.AddItem = "HOURLY2"
    'cboVGRoup.AddItem = "HOURLY3"
    'cboVGRoup.AddItem = "HOURLY4"
    'cboVGRoup.AddItem = "SALARY"
    'cboVGRoup.ListIndex = -1
    cboVGRoup.Width = 2655 'xNP_VGroup
    xNP_VGroup(0) = "1  - SALARY"
    xNP_VGroup(1) = "2  - SALARY2"
    xNP_VGroup(2) = "3  - HOURLY"
    xNP_VGroup(3) = "6  - MTHY - SALARY"
    xNP_VGroup(4) = "8  - MTHY - PER DIEM"
    xNP_VGroup(5) = "9  - Yearly Meeting"
    xNP_VGroup(6) = "10 - FIRE - SALARY no tax"
    xNP_VGroup(7) = "11 - FIRE - HOURLY"
    xNP_VGroup(8) = "12 - PT VACATION - 4%"
    xNP_VGroup(9) = "17 - HOURLY 2"
    xNP_VGroup(10) = "22 - MTHY - SALARY 1/3"
    xNP_VGroup(11) = "23 - MTHY - MTG (HALF)"
    xNP_VGroup(12) = "24 - MTHY - MTG (HALF)1/3"
    xNP_VGroup(13) = "25 - MTHY - PER DIEM 1/3"
    xNP_VGroup(14) = "37 - HOURLY - 3"
    xNP_VGroup(15) = "38 - Mayor - CAO ($60)"
    xNP_VGroup(16) = "44 - FIRE - SALARY"
    xNP_VGroup(17) = "103- FIRE - 2101"
    xNP_VGroup(18) = "104- FIRE - 2102"
    xNP_VGroup(19) = "105- FIRE - 2103"
    xNP_VGroup(20) = "114- MAYOR - CAO DUTIES"
    xNP_VGroup(21) = "115- FIRE - 2001"
    xNPVG_Cnt = 21
    For I = 0 To xNPVG_Cnt '21
        cboVGRoup.AddItem xNP_VGroup(I)
    Next
    cboVGRoup.ListIndex = -1
    'Ticket #21232 Franks 01/26/2012 - end
End Sub
Private Sub ScreenKNV() 'KN&V Ticket #21097 Franks 11/02/2011
Dim xNo As Integer
    lblTitle(19).Caption = "Pay Type"
    txtVGroup.DataField = "SH_VGROUP"
    cboVGRoup.Clear
    cboVGRoup.AddItem "SB"
    cboVGRoup.AddItem "HB"
    cboVGRoup.AddItem "H2"
    cboVGRoup.ListIndex = -1
    
    'Ticket #22952 Franks 12/10/2012 - begin
    fraSalary.Height = 1515
    lblTitle(16).Width = 2500
    lblTitle(16).Caption = "Taxable Mortgage $"
    lblTitle(18).Width = 2500
    lblTitle(18).Caption = "Non Taxable Mortgage $"
    medPremium.DataField = "SH_PREMIUM"
    medTotal.DataField = "SH_TOTAL"
    lblTitle(20).Visible = False
    cboVStep.Visible = False
    'move these field to the right using xNo
    xNo = 800
    medsalary.Left = 1670 + xNo
    medPremium.Left = 1670 + xNo
    medTotal.Left = 1670 + xNo
    lblTitle(6).Left = 3480 + xNo
    comPayPer.Left = 3840 + xNo
    lblTitle(9).Left = 5310 + xNo
    comSalScale.Left = 6300 + xNo
    lblTitle(19).Left = 3480 + xNo
    cboVGRoup.Left = 5280 + xNo
    
    fraSalary2.Top = 3490
    'Ticket #22952 Franks 12/10/2012 - end
End Sub
Private Function chkDupVGroup(xVGroup, fNew)
Dim rsLSH As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    If glbtermopen Then
        SQLQ = "SELECT SH_EMPNBR,SH_VGROUP FROM Term_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "SELECT SH_EMPNBR,SH_VGROUP FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT = 0) AND SH_EMPNBR = " & glbLEE_ID
    End If
    If fNew Then
        'nothing
    Else
        If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
            SQLQ = SQLQ & " AND NOT SH_ID = " & Data1.Recordset("SH_ID") & " "
        End If
    End If
    rsLSH.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsLSH.EOF
        If Not IsNull(rsLSH("SH_VGROUP")) Then
            If rsLSH("SH_VGROUP") = xVGroup Then
                retVal = True
            End If
        End If
        rsLSH.MoveNext
    Loop
    rsLSH.Close

    chkDupVGroup = retVal
End Function

Private Function CheckDuplCurrent(xEmpNo, xJobCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY "
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & " AND SH_JOB = '" & xJobCode & "' "
    SQLQ = SQLQ & " AND SH_CURRENT <>0 " 'current
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retVal = True
    End If
    rsTemp.Close
    CheckDuplCurrent = retVal
End Function

Private Sub FamilayDaySalaryChange(xEmpNo, xPayPeriod, xEDate, xNewRec) 'Family Day - Ticket #21152 Franks 04/01/2013
Dim rsPreSal As New ADODB.Recordset
Dim SQLQ As String
Dim xPreRate, xCurRate, xChanged
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND SH_PAYP = '" & xPayPeriod & "' "
    SQLQ = SQLQ & "AND SH_EDATE < " & Date_SQL(xEDate) & " "
    rsPreSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xPreRate = ""
    xCurRate = Val(medsalary.Text)
    If Not rsPreSal.EOF Then
        xPreRate = rsPreSal("SH_SALARY")
    End If
    rsPreSal.Close
    
    If xNewRec Then
        If Len(xPreRate) = 0 Then
            medAmtChng(1).Text = medsalary.Text
            medPercentChng(1).Text = 1
        Else '> 0
            xChanged = xCurRate - xPreRate
            medAmtChng(1).Text = xChanged
            medPercentChng(1).Text = (xChanged / xPreRate)
        End If
    Else 'changed
        If Len(xPreRate) > 0 Then 'Ticket #23779
            If Not Round(medsalary.Text, 2) = Round(xPreRate, 2) Then
                xChanged = xCurRate - xPreRate
                medAmtChng(1).Text = xChanged
                medPercentChng(1).Text = (xChanged / xPreRate)
            End If
        End If
    End If
End Sub

Private Sub WFC_Salary_US_Ben(xEmpNo) 'Ticket #23247 Franks 04/23/2013
Dim xDATE
Dim xHourWeek
    If fglbNew Then
        xDATE = dlpDate(0).Text
    Else
        xDATE = Date
    End If
    If IsNumeric(txtWHRS.Text) Then
        xHourWeek = Val(txtWHRS.Text)
    Else
        xHourWeek = 0
    End If
    Call WFC_UptUSBenByEmp(xEmpNo, CVDate(xDATE), xHourWeek)
End Sub

'Private Function Older_FollowUp_Records_Found() As Boolean
'    Dim rsFollowUp As New ADODB.Recordset
'    Dim SQLQ As String
'
'    SQLQ = "SELECT * FROM HR_FOLLOW_UP "
'    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
'    SQLQ = SQLQ & " AND EF_FREAS = 'SREV'"   'SREV, PREV, EDUC
'    SQLQ = SQLQ & " AND EF_COMPLETED <> 1"  'Not completed
'    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsFollowUp.EOF Then
'        Older_FollowUp_Records_Found = True
'    Else
'        Older_FollowUp_Records_Found = False
'    End If
'End Function

Private Sub WFCHRSoftDispValues()
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTemp

If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsCanid.EOF Then
        Exit Sub
    End If

    'Position info -------------- begin
    If Not IsNull(rsCanid("SF_SALARY")) Then
        medsalary.Text = rsCanid("SF_SALARY")
        medPercentChng(1).Text = 1 'Ticket #29438 Franks 11/07/2016
        medAmtChng(1).Text = medsalary.Text 'Ticket #29438 Franks 11/07/2016
    End If
    If Not IsNull(rsCanid("SF_STARTDATE")) Then dlpDate(0).Text = rsCanid("SF_STARTDATE")
    If Not IsNull(rsCanid("SF_SALARYFREQUENCY")) Then
        If rsCanid("SF_SALARYFREQUENCY") = "Annum" Then lblSalCode.Caption = "A"
        If rsCanid("SF_SALARYFREQUENCY") = "Hour" Then lblSalCode.Caption = "H"
        If rsCanid("SF_SALARYFREQUENCY") = "Monthly" Then lblSalCode.Caption = "M"
        If rsCanid("SF_SALARYFREQUENCY") = "Daily" Then lblSalCode.Caption = "D"
    End If
    
    rsCanid.Close
End If

End Sub

Private Function IsFirstEmpSalary(xEmpNo)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = False
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNo & " "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rs.EOF Then
        retVal = True
    End If
    rs.Close
    IsFirstEmpSalary = retVal
End Function

Public Sub imgEmailBenefit_Click()
Dim xEmail
Dim xToEmail As String

On Error GoTo Email_Err
        
        'Release 8.1
        
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetCurEmpEmail
        
        If Len(xEmail) > 0 Then
            'Ticket #18090 - begin
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            Else
                'Ticket #20317 - More Emails for everyone
                xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT", glbLEE_ID)
                If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                    xToEmail = GetComPreferEmail("EMAIL_ONBENEFIT")
                End If
            End If
            'Ticket #18090 - end
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONBENEFIT")
            If glbCompSerial = "S/N - 2382W" Then  'Samuel
            Else
                frmSendEmail.txtCC.Text = xEmail
            End If
            'frmSendEmail.txtSubject.Text = "info:HR New Benefit Notice"
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR Benefit Update Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        Else
            If Len(glbLEE_SName) = 0 Then
                MsgBox "There is no email on Status/Dates screen for employee. "
            Else
                MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            End If
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

''Private Function getWFCCurrencyIndi(xPlantCode)
''Dim rsTemp As New ADODB.Recordset
''Dim SQLQ, xStr
''Dim retVal
''
''    retVal = ""
''    If Len(xPlantCode) > 0 Then
''        SQLQ = "select * from WFC_Salary_Administration "
''        SQLQ = SQLQ & " WHERE SectionCode ='" & xPlantCode & "' "
''        SQLQ = SQLQ & " AND NOT ( CurrencyIndicator IS NULL OR CurrencyIndicator = '') "
''        SQLQ = SQLQ & "ORDER BY FiscalYear DESC"
''        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
''        If Not rsTemp.EOF Then
''            If Not IsNull(rsTemp("CurrencyIndicator")) Then
''                retVal = rsTemp("CurrencyIndicator")
''            End If
''        End If
''        rsTemp.Close
''    End If
''    getWFCCurrencyIndi = retVal
''End Function
