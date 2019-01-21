VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSLabel 
   Caption         =   "Label of Code and Date Shown"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   14100
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcDemographics 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1935
      ScaleWidth      =   3375
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   960
      Width           =   3375
      Begin Threed.SSPanel pnDemographics 
         Height          =   7335
         Left            =   0
         TabIndex        =   200
         Top             =   0
         Visible         =   0   'False
         Width           =   9855
         _Version        =   65536
         _ExtentX        =   17383
         _ExtentY        =   12938
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Email Address"
            Height          =   285
            Index           =   208
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   22
            Tag             =   "00"
            Top             =   6450
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Payroll ID"
            Height          =   285
            Index           =   198
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   1
            Tag             =   "00"
            Top             =   720
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Organization 1"
            Height          =   285
            Index           =   185
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   8
            Tag             =   "00"
            Top             =   3405
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Organization 2"
            Height          =   285
            Index           =   186
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   9
            Tag             =   "00"
            Top             =   3765
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pager Number"
            Height          =   285
            Index           =   184
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   13
            Tag             =   "00"
            Top             =   4680
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Telephone #2"
            Height          =   285
            Index           =   182
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   2
            Tag             =   "00"
            Top             =   720
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Cellular Telephone"
            Height          =   285
            Index           =   183
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   12
            Tag             =   "00"
            Top             =   4680
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Section"
            Height          =   285
            Index           =   7
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   11
            Tag             =   "00"
            Top             =   1995
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Region"
            Height          =   285
            Index           =   6
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   10
            Tag             =   "00"
            Top             =   1635
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Administered By"
            Height          =   285
            Index           =   5
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   7
            Tag             =   "00"
            Top             =   3045
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Department"
            Height          =   285
            Index           =   1
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   3
            Tag             =   "00"
            Top             =   1635
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "G/L"
            Height          =   285
            Index           =   2
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   4
            Tag             =   "00"
            Top             =   1995
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Division"
            Height          =   285
            Index           =   3
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   5
            Tag             =   "00"
            Top             =   2340
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Location"
            Height          =   285
            Index           =   4
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   6
            Tag             =   "00"
            Top             =   2685
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Driver License #"
            Height          =   285
            Index           =   24
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   14
            Tag             =   "00"
            Top             =   5040
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Type of Vehicle"
            Height          =   285
            Index           =   25
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   15
            Tag             =   "00"
            Top             =   5040
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Parking Permit #1"
            Height          =   285
            Index           =   26
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   16
            Tag             =   "00"
            Top             =   5385
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Parking Permit #2"
            Height          =   285
            Index           =   27
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   17
            Tag             =   "00"
            Top             =   5385
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "License Plate #1"
            Height          =   285
            Index           =   28
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   18
            Tag             =   "00"
            Top             =   5745
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "License Plate #2"
            Height          =   285
            Index           =   29
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   19
            Tag             =   "00"
            Top             =   5745
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Locker #"
            Height          =   285
            Index           =   30
            Left            =   2760
            MaxLength       =   30
            TabIndex        =   20
            Tag             =   "00"
            Top             =   6090
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Combination"
            Height          =   285
            Index           =   31
            Left            =   8160
            MaxLength       =   30
            TabIndex        =   21
            Tag             =   "00"
            Top             =   6090
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   208
            Left            =   2340
            TabIndex        =   690
            Top             =   6465
            Width           =   375
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Email Address"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   600
            TabIndex        =   689
            Top             =   6495
            Width           =   1425
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   198
            Left            =   2340
            TabIndex        =   662
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payroll ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   600
            TabIndex        =   661
            Top             =   765
            Width           =   675
         End
         Begin VB.Label lblBasicMiscellaneous 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Miscellaneous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   630
            Top             =   4320
            Width           =   1215
         End
         Begin VB.Label lblBasicOrganizational 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Organizational"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   629
            Top             =   1320
            Width           =   1230
         End
         Begin VB.Label lbBasicPersonal 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Personal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   628
            Top             =   435
            Width           =   750
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   186
            Left            =   2340
            TabIndex        =   627
            Top             =   3765
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   185
            Left            =   2340
            TabIndex        =   626
            Top             =   3420
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Organization 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   186
            Left            =   600
            TabIndex        =   625
            Top             =   3810
            Width           =   1020
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Organization 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   185
            Left            =   600
            TabIndex        =   624
            Top             =   3450
            Width           =   1020
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pager Number"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   5910
            TabIndex        =   623
            Top             =   4725
            Width           =   1020
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telephone #2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   6000
            TabIndex        =   622
            Top             =   765
            Width           =   1005
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
            Left            =   5910
            TabIndex        =   621
            Top             =   5085
            Width           =   1110
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   184
            Left            =   7710
            TabIndex        =   620
            Top             =   4680
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   183
            Left            =   7740
            TabIndex        =   619
            Top             =   720
            Width           =   375
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cellular Telephone"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   618
            Top             =   4725
            Width           =   1320
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   182
            Left            =   2340
            TabIndex        =   617
            Top             =   4680
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   6
            Left            =   7740
            TabIndex        =   616
            Top             =   1635
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   6000
            TabIndex        =   615
            Top             =   1680
            Width           =   510
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Division"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   600
            TabIndex        =   227
            Top             =   2385
            Width           =   555
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   600
            TabIndex        =   226
            Top             =   2730
            Width           =   615
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Administered By"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   600
            TabIndex        =   225
            Top             =   3090
            Width           =   1125
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Section"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   6000
            TabIndex        =   224
            Top             =   2040
            Width           =   540
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2340
            TabIndex        =   223
            Top             =   1635
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   2340
            TabIndex        =   222
            Top             =   2340
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   2340
            TabIndex        =   221
            Top             =   2700
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   5
            Left            =   2340
            TabIndex        =   220
            Top             =   3045
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   7
            Left            =   7740
            TabIndex        =   219
            Top             =   1995
            Width           =   375
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
            Left            =   5910
            TabIndex        =   218
            Top             =   5430
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
            Left            =   600
            TabIndex        =   217
            Top             =   5430
            Width           =   1260
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Combination"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   37
            Left            =   5910
            TabIndex        =   216
            Top             =   6135
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
            Left            =   600
            TabIndex        =   215
            Top             =   6135
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
            Left            =   5910
            TabIndex        =   214
            Top             =   5790
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
            Left            =   600
            TabIndex        =   213
            Top             =   5790
            Width           =   1200
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   21
            Left            =   7710
            TabIndex        =   212
            Top             =   6105
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   22
            Left            =   2340
            TabIndex        =   211
            Top             =   6090
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   23
            Left            =   7710
            TabIndex        =   210
            Top             =   5745
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   24
            Left            =   2340
            TabIndex        =   209
            Top             =   5745
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   25
            Left            =   7710
            TabIndex        =   208
            Top             =   5400
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   26
            Left            =   2340
            TabIndex        =   207
            Top             =   5385
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   27
            Left            =   7710
            TabIndex        =   206
            Top             =   5040
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   28
            Left            =   2340
            TabIndex        =   205
            Top             =   5040
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   2340
            TabIndex        =   204
            Top             =   1995
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "G/L"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   203
            Top             =   2040
            Width           =   285
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
            Left            =   600
            TabIndex        =   202
            Top             =   5085
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   201
            Top             =   1680
            Width           =   825
         End
      End
   End
   Begin VB.PictureBox pcJobMaster 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   9360
      ScaleHeight     =   975
      ScaleWidth      =   4215
      TabIndex        =   673
      TabStop         =   0   'False
      Top             =   9960
      Width           =   4215
      Begin Threed.SSPanel pnJobMaster 
         Height          =   975
         Left            =   480
         TabIndex        =   674
         Top             =   240
         Visible         =   0   'False
         Width           =   4815
         _Version        =   65536
         _ExtentX        =   8493
         _ExtentY        =   1720
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Job Group"
            Height          =   285
            Index           =   203
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   675
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Job Status"
            Height          =   285
            Index           =   205
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   676
            Tag             =   "00"
            Top             =   960
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Job User Defined 1"
            Height          =   285
            Index           =   206
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   677
            Tag             =   "00"
            Top             =   1440
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Job User Defined 2"
            Height          =   285
            Index           =   207
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   678
            Tag             =   "00"
            Top             =   1860
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   203
            Left            =   2520
            TabIndex        =   686
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   194
            Left            =   600
            TabIndex        =   685
            Top             =   525
            Width           =   1335
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   205
            Left            =   2520
            TabIndex        =   684
            Top             =   975
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   192
            Left            =   600
            TabIndex        =   683
            Top             =   1005
            Width           =   1215
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job User Defined 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   191
            Left            =   600
            TabIndex        =   682
            Top             =   1485
            Width           =   1365
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   206
            Left            =   2520
            TabIndex        =   681
            Top             =   1455
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Job User Defined 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   190
            Left            =   600
            TabIndex        =   680
            Top             =   1890
            Width           =   1365
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   207
            Left            =   2520
            TabIndex        =   679
            Top             =   1860
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox pcPositionMaster 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1215
      ScaleWidth      =   4335
      TabIndex        =   541
      TabStop         =   0   'False
      Top             =   8280
      Width           =   4335
      Begin Threed.SSPanel pnPositionMaster 
         Height          =   1095
         Left            =   480
         TabIndex        =   545
         Top             =   360
         Visible         =   0   'False
         Width           =   5775
         _Version        =   65536
         _ExtentX        =   10186
         _ExtentY        =   1931
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position Level"
            Height          =   285
            Index           =   204
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   189
            Tag             =   "00"
            Top             =   1280
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position User Defined 2"
            Height          =   285
            Index           =   202
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   670
            Tag             =   "00"
            Top             =   2940
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position User Defined 1"
            Height          =   285
            Index           =   201
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   667
            Tag             =   "00"
            Top             =   2520
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position Description"
            Height          =   285
            Index           =   199
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   190
            Tag             =   "00"
            Top             =   1680
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position Alternate"
            Height          =   285
            Index           =   200
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   191
            Tag             =   "00"
            Top             =   2100
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position Group"
            Height          =   285
            Index           =   32
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   187
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Position Status"
            Height          =   285
            Index           =   33
            Left            =   2940
            MaxLength       =   30
            TabIndex        =   188
            Tag             =   "00"
            Top             =   880
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   204
            Left            =   2520
            TabIndex        =   688
            Top             =   1280
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Level"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   195
            Left            =   600
            TabIndex        =   687
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   202
            Left            =   2520
            TabIndex        =   672
            Top             =   2940
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position User Defined 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   189
            Left            =   600
            TabIndex        =   671
            Top             =   2970
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   201
            Left            =   2520
            TabIndex        =   669
            Top             =   2535
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position User Defined 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   188
            Left            =   600
            TabIndex        =   668
            Top             =   2565
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Description"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   184
            Left            =   600
            TabIndex        =   666
            Top             =   1725
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   199
            Left            =   2520
            TabIndex        =   665
            Top             =   1695
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Alternate"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   187
            Left            =   600
            TabIndex        =   664
            Top             =   2145
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   200
            Left            =   2520
            TabIndex        =   663
            Top             =   2115
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   549
            Top             =   525
            Width           =   1035
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   29
            Left            =   2520
            TabIndex        =   548
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   33
            Left            =   2520
            TabIndex        =   547
            Top             =   885
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   25
            Left            =   600
            TabIndex        =   546
            Top             =   925
            Width           =   1050
         End
      End
   End
   Begin VB.PictureBox pcContEdu 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   8760
      ScaleHeight     =   1095
      ScaleWidth      =   4815
      TabIndex        =   413
      TabStop         =   0   'False
      Top             =   5520
      Width           =   4815
      Begin Threed.SSPanel frmGeneral 
         Height          =   5055
         Index           =   2
         Left            =   0
         TabIndex        =   441
         Top             =   0
         Visible         =   0   'False
         Width           =   11175
         _Version        =   65536
         _ExtentX        =   19711
         _ExtentY        =   8916
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Presenter"
            Height          =   285
            Index           =   177
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   163
            Tag             =   "00"
            Top             =   4560
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Renewal Date"
            Height          =   285
            Index           =   176
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   151
            Tag             =   "00"
            Top             =   4560
            Width           =   2355
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   7680
            Picture         =   "fxLabel.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   443
            Tag             =   "Grant All Basic"
            Top             =   5400
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Accomodation $"
            Height          =   285
            Index           =   69
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   160
            Tag             =   "00"
            Top             =   3185
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employer $"
            Height          =   285
            Index           =   68
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   159
            Tag             =   "00"
            Top             =   2845
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Expenses $"
            Height          =   285
            Index           =   67
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   158
            Tag             =   "00"
            Top             =   2505
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee $"
            Height          =   285
            Index           =   70
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   157
            Tag             =   "00"
            Top             =   2165
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Course Code"
            Height          =   285
            Index           =   77
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   140
            Tag             =   "00"
            Top             =   805
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Scheduled Date"
            Height          =   285
            Index           =   78
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   147
            Tag             =   "00"
            Top             =   3185
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Start Date"
            Height          =   285
            Index           =   79
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   148
            Tag             =   "00"
            Top             =   3525
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Course Name"
            Height          =   285
            Index           =   80
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   141
            Tag             =   "00"
            Top             =   1145
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Course Description"
            Height          =   285
            Index           =   81
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   142
            Tag             =   "00"
            Top             =   1485
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Date Completed"
            Height          =   285
            Index           =   82
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   149
            Tag             =   "00"
            Top             =   3865
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Renewal Date"
            Height          =   285
            Index           =   83
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   150
            Tag             =   "00"
            Top             =   4215
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Conducted By"
            Height          =   285
            Index           =   84
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   143
            Tag             =   "00"
            Top             =   1825
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Company Name"
            Height          =   285
            Index           =   85
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   144
            Tag             =   "00"
            Top             =   2165
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Co-Ordinated By"
            Height          =   285
            Index           =   86
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   152
            Tag             =   "00"
            Top             =   465
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Method Used"
            Height          =   285
            Index           =   87
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   153
            Tag             =   "00"
            Top             =   805
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Trainer Name"
            Height          =   285
            Index           =   88
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   145
            Tag             =   "00"
            Top             =   2505
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Results"
            Height          =   285
            Index           =   89
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   146
            Tag             =   "00"
            Top             =   2845
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Account #"
            Height          =   285
            Index           =   90
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   154
            Tag             =   "00"
            Top             =   1145
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Keyword"
            Height          =   285
            Index           =   91
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   155
            Tag             =   "00"
            Top             =   1485
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Course Hours"
            Height          =   285
            Index           =   93
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   156
            Tag             =   "00"
            Top             =   1825
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Presenter"
            Height          =   285
            Index           =   94
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   162
            Tag             =   "00"
            Top             =   4215
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Learning Material $"
            Height          =   285
            Index           =   95
            Left            =   8580
            MaxLength       =   30
            TabIndex        =   161
            Tag             =   "00"
            Top             =   3525
            Width           =   2355
         End
         Begin VB.CommandButton cmdPageRight 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            Left            =   8400
            Picture         =   "fxLabel.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   442
            Tag             =   "Grant All Basic"
            Top             =   5400
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Course Type"
            Height          =   285
            Index           =   92
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   139
            Tag             =   "00"
            Top             =   465
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   177
            Left            =   8040
            TabIndex        =   606
            Top             =   4575
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CEU Credit"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   168
            Left            =   6330
            TabIndex        =   605
            Top             =   4605
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   176
            Left            =   2370
            TabIndex        =   604
            Top             =   4575
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CEU Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   167
            Left            =   600
            TabIndex        =   603
            Top             =   4605
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   69
            Left            =   8040
            TabIndex        =   489
            Top             =   3195
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   68
            Left            =   8040
            TabIndex        =   488
            Top             =   2865
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   67
            Left            =   8040
            TabIndex        =   487
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   66
            Left            =   8040
            TabIndex        =   486
            Top             =   2175
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee $"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   58
            Left            =   6330
            TabIndex        =   485
            Top             =   2210
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Expenses $"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   59
            Left            =   6330
            TabIndex        =   484
            Top             =   2550
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employer $"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   60
            Left            =   6330
            TabIndex        =   483
            Top             =   2890
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Accommodation $"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   61
            Left            =   6330
            TabIndex        =   482
            Top             =   3230
            Width           =   1605
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Start Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   68
            Left            =   600
            TabIndex        =   481
            Top             =   3570
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Scheduled Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   69
            Left            =   600
            TabIndex        =   480
            Top             =   3230
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   70
            Left            =   600
            TabIndex        =   479
            Top             =   850
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   77
            Left            =   2370
            TabIndex        =   478
            Top             =   820
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   78
            Left            =   2370
            TabIndex        =   477
            Top             =   3200
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   79
            Left            =   2370
            TabIndex        =   476
            Top             =   3540
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Renewal Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   72
            Left            =   600
            TabIndex        =   475
            Top             =   4260
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Completed"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   73
            Left            =   600
            TabIndex        =   474
            Top             =   3910
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Description"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   74
            Left            =   600
            TabIndex        =   473
            Top             =   1530
            Width           =   1380
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   75
            Left            =   600
            TabIndex        =   472
            Top             =   1190
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   80
            Left            =   2370
            TabIndex        =   471
            Top             =   1160
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   81
            Left            =   2370
            TabIndex        =   470
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   82
            Left            =   2370
            TabIndex        =   469
            Top             =   3880
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   83
            Left            =   2370
            TabIndex        =   468
            Top             =   4230
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Method Used"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   76
            Left            =   6330
            TabIndex        =   467
            Top             =   850
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Co-ordinated By"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   77
            Left            =   6330
            TabIndex        =   466
            Top             =   510
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   78
            Left            =   600
            TabIndex        =   465
            Top             =   2210
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Conducted By"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   79
            Left            =   600
            TabIndex        =   464
            Top             =   1870
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   84
            Left            =   2370
            TabIndex        =   463
            Top             =   1840
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   85
            Left            =   2370
            TabIndex        =   462
            Top             =   2180
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   86
            Left            =   8040
            TabIndex        =   461
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   87
            Left            =   8040
            TabIndex        =   460
            Top             =   825
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Keyword"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   80
            Left            =   6330
            TabIndex        =   459
            Top             =   1530
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Account #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   81
            Left            =   6330
            TabIndex        =   458
            Top             =   1190
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Results"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   82
            Left            =   600
            TabIndex        =   457
            Top             =   2890
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Trainer Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   83
            Left            =   600
            TabIndex        =   456
            Top             =   2550
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   88
            Left            =   2370
            TabIndex        =   455
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   89
            Left            =   2370
            TabIndex        =   454
            Top             =   2860
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   90
            Left            =   8040
            TabIndex        =   453
            Top             =   1155
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   91
            Left            =   8040
            TabIndex        =   452
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Presenter"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   84
            Left            =   6330
            TabIndex        =   451
            Top             =   4260
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Hours"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   85
            Left            =   6330
            TabIndex        =   450
            Top             =   1870
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   92
            Left            =   8040
            TabIndex        =   449
            Top             =   1845
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   93
            Left            =   8040
            TabIndex        =   448
            Top             =   4230
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   95
            Left            =   8040
            TabIndex        =   447
            Top             =   3540
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Learning Material $"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   86
            Left            =   6330
            TabIndex        =   446
            Top             =   3570
            Width           =   1380
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   76
            Left            =   2370
            TabIndex        =   445
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Course Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   71
            Left            =   600
            TabIndex        =   444
            Top             =   510
            Width           =   1260
         End
      End
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   30
      SmallChange     =   4
      TabIndex        =   553
      Top             =   11830
      Width           =   14055
   End
   Begin VB.VScrollBar scrControl 
      Height          =   11625
      LargeChange     =   315
      Left            =   13820
      Max             =   2000
      SmallChange     =   315
      TabIndex        =   552
      Top             =   0
      Width           =   300
   End
   Begin Threed.SSPanel pnlLang 
      Height          =   1095
      Left            =   0
      TabIndex        =   550
      Top             =   0
      Width           =   13455
      _Version        =   65536
      _ExtentX        =   23733
      _ExtentY        =   1931
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   6
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "LB_LANG"
         DataSource      =   " "
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   0
         Tag             =   "00-Label Language - Code"
         Top             =   300
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDL1"
      End
      Begin VB.Label lblWebInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Please note that changes made to Label Master will be reflected in the web modules within one hour of your Label change."
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   566
         Top             =   720
         Width           =   8655
      End
      Begin VB.Label lblLang 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Label Language"
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
         Left            =   4320
         TabIndex        =   551
         Top             =   315
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   9000
      Top             =   11400
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   10440
      Top             =   11400
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox pcDashboard 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   8760
      ScaleHeight     =   1335
      ScaleWidth      =   4815
      TabIndex        =   591
      TabStop         =   0   'False
      Top             =   8280
      Width           =   4815
      Begin Threed.SSPanel pnDashboard 
         Height          =   5295
         Left            =   0
         TabIndex        =   592
         Top             =   360
         Width           =   13215
         _Version        =   65536
         _ExtentX        =   23310
         _ExtentY        =   9340
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.CommandButton cmdUndo 
            Appearance      =   0  'Flat
            Caption         =   "&Cancel"
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
            TabIndex        =   601
            Tag             =   "Cancel the Label Change"
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdSave 
            Appearance      =   0  'Flat
            Caption         =   "&Save"
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
            Left            =   360
            TabIndex        =   600
            Tag             =   "Save the Label Change"
            Top             =   4560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtNewItem 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   599
            Tag             =   "00-New Dashboard Item label"
            Top             =   3660
            Width           =   4935
         End
         Begin VB.TextBox txtItemCode 
            Appearance      =   0  'Flat
            DataField       =   "DB_ITEM_CODE"
            DataSource      =   "data2"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   594
            TabStop         =   0   'False
            Top             =   4080
            Visible         =   0   'False
            Width           =   615
         End
         Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
            Bindings        =   "fxLabel.frx":0884
            Height          =   2535
            Left            =   240
            OleObjectBlob   =   "fxLabel.frx":0898
            TabIndex        =   593
            Tag             =   "Dashboard Items"
            Top             =   120
            Width           =   12135
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   175
            Left            =   5640
            TabIndex        =   602
            Top             =   3660
            Width           =   375
         End
         Begin VB.Label lblItemDesc 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Item"
            DataField       =   "DB_ITEM_DESC"
            DataSource      =   "data2"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   598
            Top             =   3705
            Width           =   300
         End
         Begin VB.Label lblCategory 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            DataField       =   "DB_CATEGORY"
            DataSource      =   "data2"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1200
            TabIndex        =   597
            Top             =   3165
            Width           =   630
         End
         Begin VB.Label Label27 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Category:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   596
            Top             =   3165
            Width           =   675
         End
         Begin VB.Label Label29 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Item:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   360
            TabIndex        =   595
            Top             =   3705
            Width           =   345
         End
      End
   End
   Begin VB.PictureBox pcComments 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1215
      ScaleWidth      =   3375
      TabIndex        =   492
      TabStop         =   0   'False
      Top             =   8280
      Width           =   3375
      Begin Threed.SSPanel pnComments 
         Height          =   855
         Left            =   0
         TabIndex        =   542
         Top             =   0
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   1508
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Comments"
            Height          =   285
            Index           =   62
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   186
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments menu"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   53
            Left            =   600
            TabIndex        =   544
            Top             =   525
            Width           =   1170
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   61
            Left            =   2040
            TabIndex        =   543
            Top             =   480
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox pcUserDefined 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   360
      ScaleHeight     =   1215
      ScaleWidth      =   3375
      TabIndex        =   414
      TabStop         =   0   'False
      Top             =   6840
      Width           =   3375
      Begin Threed.SSPanel frmGeneral 
         Height          =   855
         Index           =   3
         Left            =   0
         TabIndex        =   493
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Defined Table"
            Height          =   285
            Index           =   110
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   164
            Tag             =   "00"
            Top             =   465
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Code 1"
            Height          =   285
            Index           =   111
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   165
            Tag             =   "00"
            Top             =   945
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Code 2"
            Height          =   285
            Index           =   112
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   166
            Tag             =   "00"
            Top             =   1305
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Code 3"
            Height          =   285
            Index           =   113
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   167
            Tag             =   "00"
            Top             =   1665
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Code 4"
            Height          =   285
            Index           =   114
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   168
            Tag             =   "00"
            Top             =   2025
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Code 5"
            Height          =   285
            Index           =   115
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   169
            Tag             =   "00"
            Top             =   2385
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Date 1"
            Height          =   285
            Index           =   116
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   170
            Tag             =   "00"
            Top             =   2865
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Date 2"
            Height          =   285
            Index           =   117
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   171
            Tag             =   "00"
            Top             =   3225
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Date 3"
            Height          =   285
            Index           =   118
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   172
            Tag             =   "00"
            Top             =   3585
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Date 4"
            Height          =   285
            Index           =   119
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   173
            Tag             =   "00"
            Top             =   3945
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Date 5"
            Height          =   285
            Index           =   120
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   174
            Tag             =   "00"
            Top             =   4305
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Flag 1"
            Height          =   285
            Index           =   121
            Left            =   8880
            MaxLength       =   30
            TabIndex        =   175
            Tag             =   "00"
            Top             =   2865
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Flag 2"
            Height          =   285
            Index           =   122
            Left            =   8880
            MaxLength       =   30
            TabIndex        =   176
            Tag             =   "00"
            Top             =   3240
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Flag 3"
            Height          =   285
            Index           =   123
            Left            =   8880
            MaxLength       =   30
            TabIndex        =   177
            Tag             =   "00"
            Top             =   3585
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Flag 4"
            Height          =   285
            Index           =   124
            Left            =   8880
            MaxLength       =   30
            TabIndex        =   178
            Tag             =   "00"
            Top             =   3945
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Flag 5"
            Height          =   285
            Index           =   125
            Left            =   8880
            MaxLength       =   30
            TabIndex        =   179
            Tag             =   "00"
            Top             =   4305
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Comments"
            Height          =   285
            Index           =   128
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   182
            Tag             =   "00"
            Top             =   5625
            Width           =   2355
         End
         Begin VB.CommandButton cmdPageLeft 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   3
            Left            =   7680
            Picture         =   "fxLabel.frx":631C
            Style           =   1  'Graphical
            TabIndex        =   494
            Tag             =   "Grant All Basic"
            Top             =   120
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Text 1"
            Height          =   285
            Index           =   126
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   180
            Tag             =   "00"
            Top             =   4800
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Text 2"
            Height          =   285
            Index           =   127
            Left            =   3120
            MaxLength       =   30
            TabIndex        =   181
            Tag             =   "00"
            Top             =   5160
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            AutoSize        =   -1  'True
            Caption         =   "User Defined Table menu"
            Height          =   195
            Index           =   101
            Left            =   600
            TabIndex        =   532
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblOrg 
            Caption         =   "Code 1"
            Height          =   255
            Index           =   102
            Left            =   600
            TabIndex        =   531
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Code 2"
            Height          =   255
            Index           =   103
            Left            =   600
            TabIndex        =   530
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Code 3"
            Height          =   255
            Index           =   104
            Left            =   600
            TabIndex        =   529
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Code 4"
            Height          =   255
            Index           =   105
            Left            =   600
            TabIndex        =   528
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Code 5"
            Height          =   255
            Index           =   106
            Left            =   600
            TabIndex        =   527
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Date 1"
            Height          =   255
            Index           =   107
            Left            =   600
            TabIndex        =   526
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Date 2"
            Height          =   255
            Index           =   108
            Left            =   600
            TabIndex        =   525
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Date 3"
            Height          =   255
            Index           =   109
            Left            =   600
            TabIndex        =   524
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Date 4"
            Height          =   255
            Index           =   110
            Left            =   600
            TabIndex        =   523
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Date 5"
            Height          =   255
            Index           =   111
            Left            =   600
            TabIndex        =   522
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Flag 1"
            Height          =   255
            Index           =   112
            Left            =   6360
            TabIndex        =   521
            Top             =   2880
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Flag 2"
            Height          =   255
            Index           =   113
            Left            =   6360
            TabIndex        =   520
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Flag 3"
            Height          =   255
            Index           =   114
            Left            =   6360
            TabIndex        =   519
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Flag 4"
            Height          =   255
            Index           =   115
            Left            =   6360
            TabIndex        =   518
            Top             =   3960
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Flag 5"
            Height          =   255
            Index           =   116
            Left            =   6360
            TabIndex        =   517
            Top             =   4320
            Width           =   1455
         End
         Begin VB.Label lblOrg 
            Caption         =   "Comments"
            Height          =   255
            Index           =   117
            Left            =   600
            TabIndex        =   516
            Top             =   5640
            Width           =   1455
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   109
            Left            =   2760
            TabIndex        =   515
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   110
            Left            =   2760
            TabIndex        =   514
            Top             =   960
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   111
            Left            =   2760
            TabIndex        =   513
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   112
            Left            =   2760
            TabIndex        =   512
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   113
            Left            =   2760
            TabIndex        =   511
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   114
            Left            =   2760
            TabIndex        =   510
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   115
            Left            =   2760
            TabIndex        =   509
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   116
            Left            =   2760
            TabIndex        =   508
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   117
            Left            =   2760
            TabIndex        =   507
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   118
            Left            =   2760
            TabIndex        =   506
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   119
            Left            =   2760
            TabIndex        =   505
            Top             =   4320
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   120
            Left            =   8520
            TabIndex        =   504
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   125
            Left            =   8520
            TabIndex        =   503
            Top             =   3240
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   126
            Left            =   8520
            TabIndex        =   502
            Top             =   3600
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   127
            Left            =   8520
            TabIndex        =   501
            Top             =   3960
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   128
            Left            =   8520
            TabIndex        =   500
            Top             =   4320
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   129
            Left            =   2760
            TabIndex        =   499
            Top             =   5640
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   130
            Left            =   2760
            TabIndex        =   498
            Top             =   4815
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Caption         =   "Text 1"
            Height          =   255
            Index           =   118
            Left            =   600
            TabIndex        =   497
            Top             =   4815
            Width           =   1455
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   131
            Left            =   2760
            TabIndex        =   496
            Top             =   5175
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Caption         =   "Text 2"
            Height          =   255
            Index           =   119
            Left            =   600
            TabIndex        =   495
            Top             =   5175
            Width           =   1455
         End
      End
   End
   Begin VB.PictureBox pcFollowUps 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3960
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   490
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4575
      Begin Threed.SSPanel pnFollowUps 
         Height          =   1215
         Left            =   0
         TabIndex        =   533
         Top             =   0
         Visible         =   0   'False
         Width           =   4455
         _Version        =   65536
         _ExtentX        =   7858
         _ExtentY        =   2143
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Follow-ups"
            Height          =   285
            Index           =   129
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   183
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Follow-Up"
            Height          =   285
            Index           =   130
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   184
            Tag             =   "00"
            Top             =   960
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Follow-ups menu"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   121
            Left            =   600
            TabIndex        =   537
            Top             =   525
            Width           =   1185
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   133
            Left            =   1950
            TabIndex        =   536
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Follow-Up"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   120
            Left            =   600
            TabIndex        =   535
            Top             =   1005
            Width           =   1020
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   132
            Left            =   1950
            TabIndex        =   534
            Top             =   975
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox pcCounseling 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   8760
      ScaleHeight     =   1215
      ScaleWidth      =   4695
      TabIndex        =   491
      TabStop         =   0   'False
      Top             =   6840
      Width           =   4695
      Begin Threed.SSPanel pnCounseling 
         Height          =   1095
         Left            =   0
         TabIndex        =   538
         Top             =   0
         Visible         =   0   'False
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   1931
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Counseling"
            Height          =   285
            Index           =   61
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   185
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Counselling menu"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   52
            Left            =   600
            TabIndex        =   540
            Top             =   510
            Width           =   1245
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   60
            Left            =   1950
            TabIndex        =   539
            Top             =   480
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox pcAssociations 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      ScaleHeight     =   1095
      ScaleWidth      =   4575
      TabIndex        =   412
      TabStop         =   0   'False
      Top             =   5520
      Width           =   4575
      Begin Threed.SSPanel pnAssociations 
         Height          =   975
         Left            =   0
         TabIndex        =   438
         Top             =   0
         Visible         =   0   'False
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   1720
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Associations"
            Height          =   285
            Index           =   72
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   138
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   71
            Left            =   2280
            TabIndex        =   440
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Associations menu"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   63
            Left            =   600
            TabIndex        =   439
            Top             =   525
            Width           =   1320
         End
      End
   End
   Begin VB.PictureBox pcAttendance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   3375
      TabIndex        =   402
      TabStop         =   0   'False
      Top             =   5520
      Width           =   3375
      Begin Threed.SSPanel pnAttendance 
         Height          =   975
         Left            =   0
         TabIndex        =   415
         Top             =   0
         Visible         =   0   'False
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   1720
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Machine #"
            Height          =   285
            Index           =   74
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   137
            Tag             =   "00"
            Top             =   4200
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Account Code"
            Height          =   285
            Index           =   75
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   136
            Tag             =   "00"
            Top             =   3828
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Claim #"
            Height          =   285
            Index           =   76
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   134
            Tag             =   "00"
            Top             =   3084
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Charge Code"
            Height          =   285
            Index           =   73
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   132
            Tag             =   "00"
            Top             =   2340
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "To Date"
            Height          =   285
            Index           =   152
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   128
            Tag             =   "00"
            Top             =   852
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "From Date"
            Height          =   285
            Index           =   151
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   127
            Tag             =   "00"
            Top             =   480
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Reason"
            Height          =   285
            Index           =   153
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   129
            Tag             =   "00"
            Top             =   1224
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Point"
            Height          =   285
            Index           =   157
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   135
            Tag             =   "00"
            Top             =   3456
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Shift"
            Height          =   285
            Index           =   156
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   133
            Tag             =   "00"
            Top             =   2712
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Supervisor"
            Height          =   285
            Index           =   154
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   130
            Tag             =   "00"
            Top             =   1596
            Width           =   1890
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Hours"
            Height          =   285
            Index           =   155
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   131
            Tag             =   "00"
            Top             =   1968
            Width           =   1890
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   74
            Left            =   2040
            TabIndex        =   437
            Top             =   4215
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Machine #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   65
            Left            =   600
            TabIndex        =   436
            Top             =   4245
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   75
            Left            =   2040
            TabIndex        =   435
            Top             =   3843
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Account Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   64
            Left            =   600
            TabIndex        =   434
            Top             =   3873
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   72
            Left            =   2040
            TabIndex        =   433
            Top             =   3099
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Claim #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   67
            Left            =   600
            TabIndex        =   432
            Top             =   3129
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   73
            Left            =   2070
            TabIndex        =   431
            Top             =   2355
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Charge Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   66
            Left            =   600
            TabIndex        =   430
            Top             =   2385
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "To Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   152
            Left            =   600
            TabIndex        =   429
            Top             =   897
            Width           =   585
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   164
            Left            =   2070
            TabIndex        =   428
            Top             =   867
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "From Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   151
            Left            =   600
            TabIndex        =   427
            Top             =   525
            Width           =   735
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   163
            Left            =   2070
            TabIndex        =   426
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   162
            Left            =   2070
            TabIndex        =   425
            Top             =   1239
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Reason"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   150
            Left            =   600
            TabIndex        =   424
            Top             =   1269
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Point"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   149
            Left            =   600
            TabIndex        =   423
            Top             =   3501
            Width           =   360
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   161
            Left            =   2040
            TabIndex        =   422
            Top             =   3471
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   148
            Left            =   600
            TabIndex        =   421
            Top             =   2757
            Width           =   315
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   160
            Left            =   2040
            TabIndex        =   420
            Top             =   2727
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   146
            Left            =   600
            TabIndex        =   419
            Top             =   1641
            Width           =   750
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   158
            Left            =   2070
            TabIndex        =   418
            Top             =   1611
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   159
            Left            =   2070
            TabIndex        =   417
            Top             =   1983
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   147
            Left            =   600
            TabIndex        =   416
            Top             =   2013
            Width           =   420
         End
      End
   End
   Begin VB.PictureBox pcSalaryHist 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   3960
      ScaleHeight     =   855
      ScaleWidth      =   4575
      TabIndex        =   400
      TabStop         =   0   'False
      Top             =   4440
      Width           =   4575
      Begin Threed.SSPanel pnSalaryHist 
         Height          =   855
         Left            =   0
         TabIndex        =   406
         Top             =   0
         Visible         =   0   'False
         Width           =   4095
         _Version        =   65536
         _ExtentX        =   7223
         _ExtentY        =   1508
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Comments"
            Height          =   285
            Index           =   162
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   124
            Tag             =   "00"
            Top             =   840
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pay Period"
            Height          =   285
            Index           =   23
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   123
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   153
            Left            =   600
            TabIndex        =   563
            Top             =   885
            Width           =   735
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   154
            Left            =   1995
            TabIndex        =   562
            Top             =   855
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   30
            Left            =   1995
            TabIndex        =   408
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pay Period"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   22
            Left            =   600
            TabIndex        =   407
            Top             =   525
            Width           =   765
         End
      End
   End
   Begin VB.PictureBox pcPerformanceHist 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   8760
      ScaleHeight     =   855
      ScaleWidth      =   4695
      TabIndex        =   401
      TabStop         =   0   'False
      Top             =   4440
      Width           =   4695
      Begin Threed.SSPanel pnPerformanceHist 
         Height          =   975
         Left            =   0
         TabIndex        =   409
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   1720
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Bonus $"
            Height          =   285
            Index           =   163
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   126
            Tag             =   "00"
            Top             =   840
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Performance"
            Height          =   285
            Index           =   71
            Left            =   2520
            MaxLength       =   30
            TabIndex        =   125
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Bonus $"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   154
            Left            =   600
            TabIndex        =   565
            Top             =   870
            Width           =   585
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   155
            Left            =   2160
            TabIndex        =   564
            Top             =   840
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   70
            Left            =   2160
            TabIndex        =   411
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Performance menu"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   62
            Left            =   600
            TabIndex        =   410
            Top             =   510
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox pcFlags 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   8760
      ScaleHeight     =   1095
      ScaleWidth      =   4695
      TabIndex        =   337
      TabStop         =   0   'False
      Top             =   3120
      Width           =   4695
      Begin Threed.SSPanel frmFlags 
         Height          =   1095
         Left            =   0
         TabIndex        =   359
         Top             =   0
         Visible         =   0   'False
         Width           =   3375
         _Version        =   65536
         _ExtentX        =   5953
         _ExtentY        =   1931
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 9"
            Height          =   285
            Index           =   47
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   98
            Tag             =   "00"
            Top             =   3208
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 20"
            Height          =   285
            Index           =   58
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   109
            Tag             =   "00"
            Top             =   6960
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 19"
            Height          =   285
            Index           =   57
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   108
            Tag             =   "00"
            Top             =   6618
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 18"
            Height          =   285
            Index           =   56
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   107
            Tag             =   "00"
            Top             =   6277
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 17"
            Height          =   285
            Index           =   55
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   106
            Tag             =   "00"
            Top             =   5936
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 16"
            Height          =   285
            Index           =   54
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   105
            Tag             =   "00"
            Top             =   5595
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 15"
            Height          =   285
            Index           =   53
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   104
            Tag             =   "00"
            Top             =   5254
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 14"
            Height          =   285
            Index           =   52
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   103
            Tag             =   "00"
            Top             =   4913
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 13"
            Height          =   285
            Index           =   51
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   102
            Tag             =   "00"
            Top             =   4572
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 12"
            Height          =   285
            Index           =   50
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   101
            Tag             =   "00"
            Top             =   4231
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 11"
            Height          =   285
            Index           =   49
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   100
            Tag             =   "00"
            Top             =   3890
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 10"
            Height          =   285
            Index           =   48
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   99
            Tag             =   "00"
            Top             =   3549
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 8"
            Height          =   285
            Index           =   46
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   97
            Tag             =   "00"
            Top             =   2867
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 7"
            Height          =   285
            Index           =   45
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   96
            Tag             =   "00"
            Top             =   2526
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 6"
            Height          =   285
            Index           =   44
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   95
            Tag             =   "00"
            Top             =   2185
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 5"
            Height          =   285
            Index           =   43
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   94
            Tag             =   "00"
            Top             =   1844
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 4"
            Height          =   285
            Index           =   42
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   93
            Tag             =   "00"
            Top             =   1503
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 3"
            Height          =   285
            Index           =   41
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   92
            Tag             =   "00"
            Top             =   1162
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 2"
            Height          =   285
            Index           =   40
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   91
            Tag             =   "00"
            Top             =   821
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employee Flag 1"
            Height          =   285
            Index           =   39
            Left            =   2700
            MaxLength       =   30
            TabIndex        =   90
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   57
            Left            =   2280
            TabIndex        =   399
            Top             =   5940
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   56
            Left            =   2280
            TabIndex        =   398
            Top             =   6270
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   55
            Left            =   2280
            TabIndex        =   397
            Top             =   6615
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   54
            Left            =   2280
            TabIndex        =   396
            Top             =   6975
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   53
            Left            =   2280
            TabIndex        =   395
            Top             =   5610
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 9"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   49
            Left            =   600
            TabIndex        =   394
            Top             =   3253
            Width           =   1170
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   52
            Left            =   2280
            TabIndex        =   393
            Top             =   3225
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   51
            Left            =   2280
            TabIndex        =   392
            Top             =   3570
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   50
            Left            =   2280
            TabIndex        =   391
            Top             =   5250
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   49
            Left            =   2280
            TabIndex        =   390
            Top             =   4935
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   48
            Left            =   2280
            TabIndex        =   389
            Top             =   4575
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   47
            Left            =   2280
            TabIndex        =   388
            Top             =   4230
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   46
            Left            =   2280
            TabIndex        =   387
            Top             =   3885
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   45
            Left            =   2280
            TabIndex        =   386
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   44
            Left            =   2280
            TabIndex        =   385
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   43
            Left            =   2280
            TabIndex        =   384
            Top             =   2190
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   42
            Left            =   2280
            TabIndex        =   383
            Top             =   1845
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   41
            Left            =   2280
            TabIndex        =   382
            Top             =   1515
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   40
            Left            =   2280
            TabIndex        =   381
            Top             =   1155
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   39
            Left            =   2280
            TabIndex        =   380
            Top             =   825
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   38
            Left            =   2280
            TabIndex        =   379
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 20"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   48
            Left            =   600
            TabIndex        =   378
            Top             =   7005
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 19"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   47
            Left            =   600
            TabIndex        =   377
            Top             =   6663
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 18"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   46
            Left            =   600
            TabIndex        =   376
            Top             =   6322
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 17"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   45
            Left            =   600
            TabIndex        =   375
            Top             =   5981
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 16"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   44
            Left            =   600
            TabIndex        =   374
            Top             =   5640
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 15"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   43
            Left            =   600
            TabIndex        =   373
            Top             =   5299
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 14"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   42
            Left            =   600
            TabIndex        =   372
            Top             =   4958
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 13"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   41
            Left            =   600
            TabIndex        =   371
            Top             =   4617
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 12"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   40
            Left            =   600
            TabIndex        =   370
            Top             =   4276
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 11"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   39
            Left            =   600
            TabIndex        =   369
            Top             =   3935
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 10"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   38
            Left            =   600
            TabIndex        =   368
            Top             =   3594
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 8"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   37
            Left            =   600
            TabIndex        =   367
            Top             =   2912
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 7"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   36
            Left            =   600
            TabIndex        =   366
            Top             =   2571
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 6"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   35
            Left            =   600
            TabIndex        =   365
            Top             =   2230
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 5"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   34
            Left            =   600
            TabIndex        =   364
            Top             =   1889
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   33
            Left            =   600
            TabIndex        =   363
            Top             =   1548
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   32
            Left            =   600
            TabIndex        =   362
            Top             =   1207
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   31
            Left            =   600
            TabIndex        =   361
            Top             =   866
            Width           =   1170
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   30
            Left            =   600
            TabIndex        =   360
            Top             =   525
            Width           =   1170
         End
      End
   End
   Begin VB.PictureBox pcOtherInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      ScaleHeight     =   1095
      ScaleWidth      =   4575
      TabIndex        =   336
      TabStop         =   0   'False
      Top             =   3120
      Width           =   4575
      Begin Threed.SSPanel pnOtherInfo 
         Height          =   975
         Left            =   0
         TabIndex        =   346
         Top             =   0
         Visible         =   0   'False
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   1720
         _StockProps     =   15
         Caption         =   " "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Text 3"
            Height          =   285
            Index           =   166
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   88
            Tag             =   "00"
            Top             =   3880
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Text 2"
            Height          =   285
            Index           =   165
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   87
            Tag             =   "00"
            Top             =   3480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Text 1"
            Height          =   285
            Index           =   164
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   86
            Tag             =   "00"
            Top             =   3060
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Text 4"
            Height          =   285
            Index           =   167
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   89
            Tag             =   "00"
            Top             =   4300
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Passport Number"
            Height          =   285
            Index           =   65
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   83
            Tag             =   "00"
            Top             =   1770
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Passport Country"
            Height          =   285
            Index           =   64
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   81
            Tag             =   "00"
            Top             =   915
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Citizenship"
            Height          =   285
            Index           =   63
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   80
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Visa/Work Permit #"
            Height          =   285
            Index           =   66
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   84
            Tag             =   "00"
            Top             =   2205
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Visa/Work Permit Expiration Date"
            Height          =   285
            Index           =   60
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   85
            Tag             =   "00"
            Top             =   2640
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Passport Expiration Date"
            Height          =   285
            Index           =   59
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   82
            Tag             =   "00"
            Top             =   1350
            Width           =   2355
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   166
            Left            =   3240
            TabIndex        =   574
            Top             =   3880
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Text 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   158
            Left            =   600
            TabIndex        =   573
            Top             =   3925
            Width           =   885
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   165
            Left            =   3240
            TabIndex        =   572
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Text 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   157
            Left            =   600
            TabIndex        =   571
            Top             =   3525
            Width           =   885
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   157
            Left            =   3240
            TabIndex        =   570
            Top             =   3060
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Text 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   156
            Left            =   600
            TabIndex        =   569
            Top             =   3105
            Width           =   885
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   156
            Left            =   3240
            TabIndex        =   568
            Top             =   4300
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Text 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   155
            Left            =   600
            TabIndex        =   567
            Top             =   4345
            Width           =   885
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Passport Number"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   57
            Left            =   600
            TabIndex        =   358
            Top             =   1815
            Width           =   1215
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   65
            Left            =   3240
            TabIndex        =   357
            Top             =   1770
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Passport Country"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   56
            Left            =   600
            TabIndex        =   356
            Top             =   960
            Width           =   1200
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   64
            Left            =   3240
            TabIndex        =   355
            Top             =   915
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   55
            Left            =   600
            TabIndex        =   354
            Top             =   525
            Width           =   750
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   63
            Left            =   3240
            TabIndex        =   353
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Visa/Work Permit #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   54
            Left            =   600
            TabIndex        =   352
            Top             =   2250
            Width           =   1395
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   62
            Left            =   3240
            TabIndex        =   351
            Top             =   2205
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Visa/Work Permit Expiration Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   50
            Left            =   600
            TabIndex        =   350
            Top             =   2685
            Width           =   2370
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   58
            Left            =   3240
            TabIndex        =   349
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   59
            Left            =   3240
            TabIndex        =   348
            Top             =   1350
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Passport Expiration Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   51
            Left            =   600
            TabIndex        =   347
            Top             =   1395
            Width           =   1740
         End
      End
   End
   Begin VB.PictureBox pcBanking 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   360
      ScaleHeight     =   1095
      ScaleWidth      =   3375
      TabIndex        =   230
      TabStop         =   0   'False
      Top             =   3120
      Width           =   3375
      Begin Threed.SSPanel pnBanking 
         Height          =   1335
         Left            =   0
         TabIndex        =   339
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   2355
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Vadim Field 1"
            Height          =   285
            Index           =   37
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   78
            Tag             =   "00"
            Top             =   890
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Vadim Field 2"
            Height          =   285
            Index           =   38
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   79
            Tag             =   "00"
            Top             =   1300
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Supervisor Code"
            Height          =   285
            Index           =   36
            Left            =   2400
            MaxLength       =   30
            TabIndex        =   77
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vadim Field 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   26
            Left            =   600
            TabIndex        =   345
            Top             =   935
            Width           =   945
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   34
            Left            =   2020
            TabIndex        =   344
            Top             =   905
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Vadim Field 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   28
            Left            =   600
            TabIndex        =   343
            Top             =   1345
            Width           =   945
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   36
            Left            =   2020
            TabIndex        =   342
            Top             =   1315
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   37
            Left            =   2020
            TabIndex        =   341
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Supervisor Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   29
            Left            =   600
            TabIndex        =   340
            Top             =   525
            Width           =   1170
         End
      End
   End
   Begin VB.PictureBox pcDependents 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8760
      ScaleHeight     =   1575
      ScaleWidth      =   4695
      TabIndex        =   229
      TabStop         =   0   'False
      Top             =   1320
      Width           =   4695
      Begin Threed.SSPanel pnDependents 
         Height          =   1215
         Left            =   0
         TabIndex        =   317
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   2143
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Text 4"
            Height          =   285
            Index           =   171
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   76
            Tag             =   "00"
            Top             =   2520
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Text 2"
            Height          =   285
            Index           =   169
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   75
            Tag             =   "00"
            Top             =   2100
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Text 1"
            Height          =   285
            Index           =   168
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   68
            Tag             =   "00"
            Top             =   2100
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Text 3"
            Height          =   285
            Index           =   170
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   69
            Tag             =   "00"
            Top             =   2520
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Comment"
            Height          =   285
            Index           =   108
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   70
            Tag             =   "00"
            Top             =   2940
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Benefit End Date"
            Height          =   285
            Index           =   107
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   67
            Tag             =   "00"
            Top             =   1695
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "COB Other"
            Height          =   285
            Index           =   106
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   73
            Tag             =   "00"
            Top             =   1290
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Benefit Eligible Date"
            Height          =   285
            Index           =   105
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   66
            Tag             =   "00"
            Top             =   1290
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "COB Medical"
            Height          =   285
            Index           =   104
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   72
            Tag             =   "00"
            Top             =   885
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Smoker"
            Height          =   285
            Index           =   103
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   65
            Tag             =   "00"
            Top             =   885
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "COB Dental"
            Height          =   285
            Index           =   102
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   71
            Tag             =   "00"
            Top             =   480
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Status"
            Height          =   285
            Index           =   101
            Left            =   2880
            MaxLength       =   50
            TabIndex        =   64
            Tag             =   "00"
            Top             =   480
            Width           =   1875
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Dependent Number"
            Height          =   285
            Index           =   109
            Left            =   7440
            MaxLength       =   50
            TabIndex        =   74
            Tag             =   "00"
            Top             =   1695
            Width           =   1875
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   170
            Left            =   7080
            TabIndex        =   582
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Text 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   162
            Left            =   5040
            TabIndex        =   581
            Top             =   2565
            Width           =   1290
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   169
            Left            =   7080
            TabIndex        =   580
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Text 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   161
            Left            =   5040
            TabIndex        =   579
            Top             =   2145
            Width           =   1290
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Text 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   160
            Left            =   600
            TabIndex        =   578
            Top             =   2145
            Width           =   1290
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   168
            Left            =   2520
            TabIndex        =   577
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Text 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   159
            Left            =   600
            TabIndex        =   576
            Top             =   2565
            Width           =   1290
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   167
            Left            =   2550
            TabIndex        =   575
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   108
            Left            =   2550
            TabIndex        =   335
            Top             =   1290
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Eligible Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   100
            Left            =   600
            TabIndex        =   334
            Top             =   1335
            Width           =   1425
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   107
            Left            =   2550
            TabIndex        =   333
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit End Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   99
            Left            =   600
            TabIndex        =   332
            Top             =   1740
            Width           =   1215
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "COB Other"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   98
            Left            =   5040
            TabIndex        =   331
            Top             =   1335
            Width           =   1845
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   106
            Left            =   7080
            TabIndex        =   330
            Top             =   1290
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Comment"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   97
            Left            =   600
            TabIndex        =   329
            Top             =   2985
            Width           =   1500
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   105
            Left            =   2520
            TabIndex        =   328
            Top             =   2940
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   104
            Left            =   7080
            TabIndex        =   327
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Number"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   96
            Left            =   5040
            TabIndex        =   326
            Top             =   1740
            Width           =   1395
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   103
            Left            =   2550
            TabIndex        =   325
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Status"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   95
            Left            =   600
            TabIndex        =   324
            Top             =   525
            Width           =   1290
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   102
            Left            =   2550
            TabIndex        =   323
            Top             =   885
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Dependent Smoker"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   94
            Left            =   600
            TabIndex        =   322
            Top             =   930
            Width           =   1380
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "COB Dental"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   93
            Left            =   5040
            TabIndex        =   321
            Top             =   525
            Width           =   1920
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   101
            Left            =   7080
            TabIndex        =   320
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "COB Medical"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   92
            Left            =   5040
            TabIndex        =   319
            Top             =   930
            Width           =   1890
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   100
            Left            =   7080
            TabIndex        =   318
            Top             =   885
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox pcProvince 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   360
      ScaleHeight     =   1935
      ScaleWidth      =   3375
      TabIndex        =   631
      TabStop         =   0   'False
      Top             =   9720
      Width           =   3375
      Begin Threed.SSPanel pnProvince 
         Height          =   2535
         Left            =   0
         TabIndex        =   632
         Top             =   0
         Visible         =   0   'False
         Width           =   3255
         _Version        =   65536
         _ExtentX        =   5741
         _ExtentY        =   4471
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. Text3"
            Height          =   285
            Index           =   193
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   198
            Tag             =   "00"
            Top             =   2640
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. Text2"
            Height          =   285
            Index           =   192
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   197
            Tag             =   "00"
            Top             =   2280
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. Text1"
            Height          =   285
            Index           =   191
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   196
            Tag             =   "00"
            Top             =   1920
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. Num2"
            Height          =   285
            Index           =   190
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   195
            Tag             =   "00"
            Top             =   1560
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. Num1"
            Height          =   285
            Index           =   189
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   194
            Tag             =   "00"
            Top             =   1200
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. #"
            Height          =   285
            Index           =   188
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   193
            Tag             =   "00"
            Top             =   840
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Prov. Name"
            Height          =   285
            Index           =   187
            Left            =   2340
            MaxLength       =   30
            TabIndex        =   192
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. Text3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   179
            Left            =   600
            TabIndex        =   646
            Top             =   2685
            Width           =   825
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   193
            Left            =   1920
            TabIndex        =   645
            Top             =   2640
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. Text2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   178
            Left            =   600
            TabIndex        =   644
            Top             =   2325
            Width           =   825
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   192
            Left            =   1920
            TabIndex        =   643
            Top             =   2275
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. Text1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   177
            Left            =   600
            TabIndex        =   642
            Top             =   1965
            Width           =   825
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   191
            Left            =   1920
            TabIndex        =   641
            Top             =   1913
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. Num2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   176
            Left            =   600
            TabIndex        =   640
            Top             =   1605
            Width           =   840
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   190
            Left            =   1920
            TabIndex        =   639
            Top             =   1551
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. Num1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   175
            Left            =   600
            TabIndex        =   638
            Top             =   1245
            Width           =   840
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   189
            Left            =   1920
            TabIndex        =   637
            Top             =   1189
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prov. #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   174
            Left            =   600
            TabIndex        =   636
            Top             =   885
            Width           =   525
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   188
            Left            =   1920
            TabIndex        =   635
            Top             =   827
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   187
            Left            =   1920
            TabIndex        =   634
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   173
            Left            =   600
            TabIndex        =   633
            Top             =   525
            Width           =   420
         End
      End
   End
   Begin VB.PictureBox pcStatusDates 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   3960
      ScaleHeight     =   2055
      ScaleWidth      =   4575
      TabIndex        =   228
      TabStop         =   0   'False
      Top             =   840
      Width           =   4575
      Begin Threed.SSPanel pnStatusDates 
         Height          =   1890
         Left            =   0
         TabIndex        =   231
         Top             =   0
         Visible         =   0   'False
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   3334
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Category"
            Height          =   285
            Index           =   9
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   24
            Tag             =   "00"
            Top             =   807
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Number 2"
            Height          =   285
            Index           =   99
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   34
            Tag             =   "00"
            Top             =   1455
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Date"
            Height          =   285
            Index           =   100
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   35
            Tag             =   "00"
            Top             =   1785
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 10"
            Height          =   285
            Index           =   146
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   63
            Tag             =   "00"
            Top             =   9630
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 9"
            Height          =   285
            Index           =   145
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   62
            Tag             =   "00"
            Top             =   9285
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 5"
            Height          =   285
            Index           =   141
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   58
            Tag             =   "00"
            Top             =   9630
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 4"
            Height          =   285
            Index           =   140
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   57
            Tag             =   "00"
            Top             =   9285
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 8"
            Height          =   285
            Index           =   144
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   61
            Tag             =   "00"
            Top             =   8955
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 7"
            Height          =   285
            Index           =   143
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   60
            Tag             =   "00"
            Top             =   8610
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 6"
            Height          =   285
            Index           =   142
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   59
            Tag             =   "00"
            Top             =   8280
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 3"
            Height          =   285
            Index           =   139
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   56
            Tag             =   "00"
            Top             =   8955
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 2"
            Height          =   285
            Index           =   138
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   55
            Tag             =   "00"
            Top             =   8610
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Other Date 1"
            Height          =   285
            Index           =   137
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   54
            Tag             =   "00"
            Top             =   8280
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pension Date 1"
            Height          =   285
            Index           =   131
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   48
            Tag             =   "00"
            Top             =   5940
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pension Date 3"
            Height          =   285
            Index           =   133
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   50
            Tag             =   "00"
            Top             =   6585
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pension Date 4"
            Height          =   285
            Index           =   134
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   51
            Tag             =   "00"
            Top             =   6915
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pension Date 5"
            Height          =   285
            Index           =   135
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   52
            Tag             =   "00"
            Top             =   7230
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Normal Retirement"
            Height          =   285
            Index           =   20
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   46
            Tag             =   "00"
            Top             =   6585
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Latest Retirement"
            Height          =   285
            Index           =   21
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   47
            Tag             =   "00"
            Top             =   6915
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Eligibility"
            Height          =   285
            Index           =   18
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   44
            Tag             =   "00"
            Top             =   5940
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Defined"
            Height          =   285
            Index           =   17
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   43
            Tag             =   "00"
            Top             =   4950
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "OMERS Date"
            Height          =   285
            Index           =   16
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   42
            Tag             =   "00"
            Top             =   4620
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Last Day"
            Height          =   285
            Index           =   15
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   41
            Tag             =   "00"
            Top             =   4290
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "First Day"
            Height          =   285
            Index           =   14
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   40
            Tag             =   "00"
            Top             =   3960
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Union Date"
            Height          =   285
            Index           =   13
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   39
            Tag             =   "00"
            Top             =   4950
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Last Hire"
            Height          =   285
            Index           =   12
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   38
            Tag             =   "00"
            Top             =   4620
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Seniority"
            Height          =   285
            Index           =   11
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   37
            Tag             =   "00"
            Top             =   4290
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Original Hire"
            Height          =   285
            Index           =   10
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   36
            Tag             =   "00"
            Top             =   3960
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Text 1"
            Height          =   285
            Index           =   96
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   29
            Tag             =   "00"
            Top             =   2442
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Number 1"
            Height          =   285
            Index           =   98
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   30
            Tag             =   "00"
            Top             =   2775
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Hire Code"
            Height          =   285
            Index           =   22
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   28
            Tag             =   "00"
            Top             =   2115
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Union"
            Height          =   285
            Index           =   8
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   25
            Tag             =   "00"
            Top             =   1134
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Employment Type"
            Height          =   285
            Index           =   147
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   23
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Internal Phone Extension"
            Height          =   285
            Index           =   149
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   26
            Tag             =   "00"
            Top             =   1461
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Email Address"
            Height          =   285
            Index           =   150
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   27
            Tag             =   "00"
            Top             =   1788
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pension Date 6"
            Height          =   285
            Index           =   136
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   53
            Tag             =   "00"
            Top             =   7560
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Earliest Retirement"
            Height          =   285
            Index           =   19
            Left            =   2910
            MaxLength       =   30
            TabIndex        =   45
            Tag             =   "00"
            Top             =   6270
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Pension Date 2"
            Height          =   285
            Index           =   132
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   49
            Tag             =   "00"
            Top             =   6270
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "User Text 2"
            Height          =   285
            Index           =   97
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   33
            Tag             =   "00"
            Top             =   1125
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Benefit Group"
            Height          =   285
            Index           =   148
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   31
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Salary Distribution"
            Height          =   285
            Index           =   35
            Left            =   7920
            MaxLength       =   30
            TabIndex        =   32
            Tag             =   "00"
            Top             =   807
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   141
            Left            =   600
            TabIndex        =   316
            Top             =   1833
            Width           =   990
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   153
            Left            =   2460
            TabIndex        =   315
            Top             =   1803
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   152
            Left            =   2460
            TabIndex        =   314
            Top             =   1149
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Internal Phone Extension"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   140
            Left            =   600
            TabIndex        =   313
            Top             =   1506
            Width           =   1770
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Employment Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   139
            Left            =   600
            TabIndex        =   312
            Top             =   525
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   149
            Left            =   2460
            TabIndex        =   311
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Union"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   600
            TabIndex        =   310
            Top             =   1179
            Width           =   420
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   2460
            TabIndex        =   309
            Top             =   1476
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2460
            TabIndex        =   308
            Top             =   822
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   600
            TabIndex        =   307
            Top             =   852
            Width           =   630
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   31
            Left            =   2460
            TabIndex        =   306
            Top             =   2115
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hire Code"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   23
            Left            =   600
            TabIndex        =   305
            Top             =   2160
            Width           =   705
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Salary Distribution"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   27
            Left            =   5760
            TabIndex        =   304
            Top             =   852
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   35
            Left            =   7470
            TabIndex        =   303
            Top             =   822
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Benefit Group"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   136
            Left            =   5760
            TabIndex        =   302
            Top             =   525
            Width           =   1095
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   147
            Left            =   7470
            TabIndex        =   301
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   99
            Left            =   7470
            TabIndex        =   300
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   91
            Left            =   5760
            TabIndex        =   299
            Top             =   1830
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   98
            Left            =   7470
            TabIndex        =   298
            Top             =   1470
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   97
            Left            =   2460
            TabIndex        =   297
            Top             =   2790
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User Number 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   90
            Left            =   600
            TabIndex        =   296
            Top             =   2820
            Width           =   1500
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User Number 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   89
            Left            =   5760
            TabIndex        =   295
            Top             =   1500
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   96
            Left            =   7470
            TabIndex        =   294
            Top             =   1140
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   94
            Left            =   2460
            TabIndex        =   293
            Top             =   2457
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User Text 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   88
            Left            =   600
            TabIndex        =   292
            Top             =   2487
            Width           =   1260
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User Text 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   87
            Left            =   5760
            TabIndex        =   291
            Top             =   1170
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   148
            Left            =   2460
            TabIndex        =   290
            Top             =   9300
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   146
            Left            =   2460
            TabIndex        =   289
            Top             =   9630
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 10"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   138
            Left            =   5760
            TabIndex        =   288
            Top             =   9675
            Width           =   1005
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 9"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   137
            Left            =   5760
            TabIndex        =   287
            Top             =   9330
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 5"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   135
            Left            =   600
            TabIndex        =   286
            Top             =   9675
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   134
            Left            =   600
            TabIndex        =   285
            Top             =   9330
            Width           =   915
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   142
            Left            =   2460
            TabIndex        =   284
            Top             =   8295
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   141
            Left            =   2460
            TabIndex        =   283
            Top             =   8955
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   140
            Left            =   2460
            TabIndex        =   282
            Top             =   8610
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 8"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   133
            Left            =   5760
            TabIndex        =   281
            Top             =   9000
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 7"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   132
            Left            =   5760
            TabIndex        =   280
            Top             =   8655
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 6"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   131
            Left            =   5760
            TabIndex        =   279
            Top             =   8325
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   130
            Left            =   600
            TabIndex        =   278
            Top             =   9000
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   129
            Left            =   600
            TabIndex        =   277
            Top             =   8655
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Other Date 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   128
            Left            =   600
            TabIndex        =   276
            Top             =   8325
            Width           =   915
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   127
            Left            =   5760
            TabIndex        =   275
            Top             =   5985
            Width           =   1095
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   126
            Left            =   5760
            TabIndex        =   274
            Top             =   6315
            Width           =   1095
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   125
            Left            =   5760
            TabIndex        =   273
            Top             =   6630
            Width           =   1095
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   124
            Left            =   5760
            TabIndex        =   272
            Top             =   6960
            Width           =   1095
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   139
            Left            =   7470
            TabIndex        =   271
            Top             =   6270
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   138
            Left            =   7470
            TabIndex        =   270
            Top             =   6585
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   137
            Left            =   7470
            TabIndex        =   269
            Top             =   5955
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   134
            Left            =   7470
            TabIndex        =   268
            Top             =   6915
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Normal Retirement"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   20
            Left            =   600
            TabIndex        =   267
            Top             =   6630
            Width           =   1305
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   19
            Left            =   7470
            TabIndex        =   266
            Top             =   4950
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   18
            Left            =   2460
            TabIndex        =   265
            Top             =   5940
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   17
            Left            =   2460
            TabIndex        =   264
            Top             =   6270
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   2460
            TabIndex        =   263
            Top             =   6600
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   15
            Left            =   2460
            TabIndex        =   262
            Top             =   6915
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   7470
            TabIndex        =   261
            Top             =   4635
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   20
            Left            =   2460
            TabIndex        =   260
            Top             =   3975
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   13
            Left            =   7470
            TabIndex        =   259
            Top             =   4290
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   7470
            TabIndex        =   258
            Top             =   3975
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   11
            Left            =   2460
            TabIndex        =   257
            Top             =   4950
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   10
            Left            =   2460
            TabIndex        =   256
            Top             =   4620
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   9
            Left            =   2460
            TabIndex        =   255
            Top             =   4290
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Latest Retirement"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   21
            Left            =   600
            TabIndex        =   254
            Top             =   6960
            Width           =   1245
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Earliest Retirement"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   19
            Left            =   600
            TabIndex        =   253
            Top             =   6315
            Width           =   1320
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Eligibility"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   18
            Left            =   600
            TabIndex        =   252
            Top             =   5985
            Width           =   585
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User Defined"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   17
            Left            =   5760
            TabIndex        =   251
            Top             =   4995
            Width           =   930
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OMERS Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   16
            Left            =   5760
            TabIndex        =   250
            Top             =   4665
            Width           =   975
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Day"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   15
            Left            =   5760
            TabIndex        =   249
            Top             =   4335
            Width           =   630
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "First Day"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   14
            Left            =   5760
            TabIndex        =   248
            Top             =   4005
            Width           =   615
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Union Date"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   13
            Left            =   600
            TabIndex        =   247
            Top             =   4995
            Width           =   810
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Last Hire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   12
            Left            =   600
            TabIndex        =   246
            Top             =   4665
            Width           =   630
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Seniority"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   600
            TabIndex        =   245
            Top             =   4335
            Width           =   600
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Original Hire"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   600
            TabIndex        =   244
            Top             =   4005
            Width           =   855
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 6"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   122
            Left            =   5760
            TabIndex        =   243
            Top             =   7605
            Width           =   1095
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   143
            Left            =   7470
            TabIndex        =   242
            Top             =   8955
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   144
            Left            =   7470
            TabIndex        =   241
            Top             =   8610
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   145
            Left            =   7470
            TabIndex        =   240
            Top             =   8280
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   150
            Left            =   7470
            TabIndex        =   239
            Top             =   9600
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   151
            Left            =   7470
            TabIndex        =   238
            Top             =   9285
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   136
            Left            =   7470
            TabIndex        =   237
            Top             =   7560
            Width           =   375
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Employment Dates"
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
            Index           =   2
            Left            =   360
            TabIndex        =   236
            Top             =   3600
            Width           =   1575
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Pension Dates"
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
            Index           =   4
            Left            =   360
            TabIndex        =   235
            Top             =   5640
            Width           =   1245
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Other Dates"
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
            Index           =   5
            Left            =   360
            TabIndex        =   234
            Top             =   7920
            Width           =   1035
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   135
            Left            =   7470
            TabIndex        =   233
            Top             =   7230
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pension Date 5"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   123
            Left            =   5760
            TabIndex        =   232
            Top             =   7275
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox pcPositionHist 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   360
      ScaleHeight     =   855
      ScaleWidth      =   3375
      TabIndex        =   338
      TabStop         =   0   'False
      Top             =   4440
      Width           =   3375
      Begin Threed.SSPanel pnPositionHist 
         Height          =   855
         Left            =   0
         TabIndex        =   403
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   1508
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Hours/Pay Period"
            Height          =   285
            Index           =   180
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   118
            Tag             =   "00"
            Top             =   3360
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Hours/Week"
            Height          =   285
            Index           =   179
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   117
            Tag             =   "00"
            Top             =   3000
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Hours/Day"
            Height          =   285
            Index           =   178
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   116
            Tag             =   "00"
            Top             =   2640
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "FTE Hours/Year"
            Height          =   285
            Index           =   181
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   120
            Tag             =   "00"
            Top             =   4080
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Rept. Authority 4"
            Height          =   285
            Index           =   175
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   115
            Tag             =   "00"
            Top             =   2280
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Rept. Authority 3"
            Height          =   285
            Index           =   174
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   114
            Tag             =   "00"
            Top             =   1920
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Rept. Authority 2"
            Height          =   285
            Index           =   173
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   113
            Tag             =   "00"
            Top             =   1560
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Rept. Authority 1"
            Height          =   285
            Index           =   172
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   112
            Tag             =   "00"
            Top             =   1200
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Notes 2"
            Height          =   285
            Index           =   161
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   122
            Tag             =   "00"
            Top             =   4800
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Notes 1"
            Height          =   285
            Index           =   160
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   121
            Tag             =   "00"
            Top             =   4440
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Shift"
            Height          =   285
            Index           =   159
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   119
            Tag             =   "00"
            Top             =   3720
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Acting Position"
            Height          =   285
            Index           =   158
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   111
            Tag             =   "00"
            Top             =   840
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Salary Category"
            Height          =   285
            Index           =   34
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   110
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "FTE Hours/Year"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   172
            Left            =   600
            TabIndex        =   614
            Top             =   4125
            Width           =   1170
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   181
            Left            =   2280
            TabIndex        =   613
            Top             =   4095
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Pay Period"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   171
            Left            =   600
            TabIndex        =   612
            Top             =   3405
            Width           =   1260
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   180
            Left            =   2280
            TabIndex        =   611
            Top             =   3375
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Week"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   170
            Left            =   600
            TabIndex        =   610
            Top             =   3045
            Width           =   930
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   179
            Left            =   2280
            TabIndex        =   609
            Top             =   3015
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Day"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   169
            Left            =   600
            TabIndex        =   608
            Top             =   2685
            Width           =   780
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   178
            Left            =   2280
            TabIndex        =   607
            Top             =   2655
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   174
            Left            =   2280
            TabIndex        =   590
            Top             =   2295
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rept. Authority 4"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   166
            Left            =   600
            TabIndex        =   589
            Top             =   2325
            Width           =   1185
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   173
            Left            =   2280
            TabIndex        =   588
            Top             =   1935
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rept. Authority 3"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   165
            Left            =   600
            TabIndex        =   587
            Top             =   1965
            Width           =   1185
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   172
            Left            =   2280
            TabIndex        =   586
            Top             =   1575
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rept. Authority 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   164
            Left            =   600
            TabIndex        =   585
            Top             =   1605
            Width           =   1185
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   171
            Left            =   2280
            TabIndex        =   584
            Top             =   1215
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Rept. Authority 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   163
            Left            =   600
            TabIndex        =   583
            Top             =   1245
            Width           =   1185
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notes 2"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   145
            Left            =   600
            TabIndex        =   561
            Top             =   4845
            Width           =   555
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   124
            Left            =   2280
            TabIndex        =   560
            Top             =   4815
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Notes 1"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   144
            Left            =   600
            TabIndex        =   559
            Top             =   4485
            Width           =   555
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   123
            Left            =   2280
            TabIndex        =   558
            Top             =   4455
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Shift"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   143
            Left            =   600
            TabIndex        =   557
            Top             =   3765
            Width           =   315
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   122
            Left            =   2280
            TabIndex        =   556
            Top             =   3735
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Acting Position"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   142
            Left            =   600
            TabIndex        =   555
            Top             =   885
            Width           =   1050
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   121
            Left            =   2280
            TabIndex        =   554
            Top             =   855
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   32
            Left            =   2280
            TabIndex        =   405
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Grid Category"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   24
            Left            =   600
            TabIndex        =   404
            Top             =   525
            Width           =   960
         End
      End
   End
   Begin VB.PictureBox pcAddPayrollIDData 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   3960
      ScaleHeight     =   2295
      ScaleWidth      =   5175
      TabIndex        =   647
      TabStop         =   0   'False
      Top             =   9720
      Width           =   5175
      Begin Threed.SSPanel pnAddPayrollIDData 
         Height          =   1695
         Left            =   0
         TabIndex        =   648
         Top             =   0
         Visible         =   0   'False
         Width           =   4695
         _Version        =   65536
         _ExtentX        =   8281
         _ExtentY        =   2990
         _StockProps     =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         Alignment       =   0
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "ADP Department"
            Height          =   285
            Index           =   197
            Left            =   3480
            MaxLength       =   30
            TabIndex        =   652
            Tag             =   "00"
            Top             =   1680
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "ADP GL #"
            Height          =   285
            Index           =   196
            Left            =   3480
            MaxLength       =   30
            TabIndex        =   651
            Tag             =   "00"
            Top             =   1320
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "ADP Branch #"
            Height          =   285
            Index           =   195
            Left            =   3480
            MaxLength       =   30
            TabIndex        =   650
            Tag             =   "00"
            Top             =   960
            Width           =   2355
         End
         Begin VB.TextBox txtNew 
            Appearance      =   0  'Flat
            DataField       =   "Additional Payroll ID Data"
            Height          =   285
            Index           =   194
            Left            =   3480
            MaxLength       =   30
            TabIndex        =   649
            Tag             =   "00"
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ADP Department"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   183
            Left            =   600
            TabIndex        =   660
            Top             =   1725
            Width           =   1200
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   197
            Left            =   3060
            TabIndex        =   659
            Top             =   1695
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ADP GL #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   182
            Left            =   600
            TabIndex        =   658
            Top             =   1365
            Width           =   1140
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   196
            Left            =   3060
            TabIndex        =   657
            Top             =   1335
            Width           =   375
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   195
            Left            =   3060
            TabIndex        =   656
            Top             =   975
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ADP Branch #"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   181
            Left            =   600
            TabIndex        =   655
            Top             =   1005
            Width           =   1140
         End
         Begin VB.Label lblArrow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "-->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   194
            Left            =   3060
            TabIndex        =   654
            Top             =   495
            Width           =   375
         End
         Begin VB.Label lblOrg 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Additional Payroll ID Data menu"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   180
            Left            =   600
            TabIndex        =   653
            Top             =   525
            Width           =   2235
         End
      End
   End
End
Attribute VB_Name = "frmSLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fglbNew As Boolean
Dim ONewItem As String

Private Function chkLabel()
    chkLabel = False
    
    If Len(Trim(clpCode(0).Text)) = 0 Then
        MsgBox "Label Language cannot be left blank"
        clpCode(0).SetFocus
        Exit Function
    End If
    If Not clpCode(0).ListChecker Then Exit Function
    
    chkLabel = True
End Function

Sub cmdCancel_Click()
Dim X
Dim bmk As Variant

'Ticket #22825
'If it's Dashboard Item - then check for values changed & not saved differently
If Me.Caption = "Label - Dashboard Setup screen" Then
    If data2.Recordset.EOF And data2.Recordset.BOF Then
        bmk = 0
    Else
        bmk = data2.Recordset.Bookmark
    End If
    
    'Refresh Data2
    data2.Refresh
    If Not bmk = 0 Then
        data2.Recordset.Bookmark = bmk
    End If
    
    'Store original values
    ONewItem = txtNewItem.Text
End If

X = EERetrieve

Call ST_UPD_MODE(True)

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

'Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
'Call SET_UP_MODE
Call ST_UPD_MODE(True)
End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X
'If Not (txtNew(61) = "Counseling" Or txtNew(61) = "Counselling") Then
'    MsgBox "'Counseling' only can be changed to 'Counselling'"
'    Exit Sub
'End If

If Not chkLabel Then Exit Sub

'Ticket #22825
'If it's Dashboard Item - then check for values changed & not saved differently
If Me.Caption = "Label - Dashboard Setup screen" Then
    If ONewItem <> txtNewItem.Text Then
        'Save the changes
        Call cmdSave_Click
        
        GoTo Contd
    End If
End If

Data1.RecordSource = "SELECT * FROM HRLABEL WHERE LB_LANG = '" & clpCode(0).Text & "'"
Data1.Refresh

For X = 1 To glbLabels(2).count
    glbLabels(2).Remove 1
Next

Data1.Refresh
For X = 1 To glbLabels(1).count
    With Data1.Recordset
        If Not .EOF Then .MoveFirst
        .Find "LB_ORG='" & glbLabels(1)(X) & "'"
        If .EOF Then
            .AddNew
        End If
        !LB_COMPNO = "001"
        If txtNew(X) = "" And (X < 39 Or X > 58) Then
            !LB_NEW = glbLabels(1)(X)
        Else
            !LB_NEW = txtNew(X)
        End If
        !LB_ORG = glbLabels(1)(X)
        !LB_LANG = clpCode(0).Text
        !LB_LANG_TABL = "LBLG"
        
        !LB_LDATE = Date
        !LB_LTIME = Time$
        !LB_LUSER = glbUserID
        
        .Update
        glbLabels(2).Add Trim(txtNew(X)), Trim(glbLabels(1)(X))
    End With
Next
glbLabLang = clpCode(0).Text
Data1.Refresh

Contd:
fglbNew = False

'Call SET_UP_MODE   'Causing an issue when the menu label changes
Dim UpdateState As UpdateStateEnum
UpdateState = OPENING
Call set_Buttons(UpdateState)

'Call ST_UPD_MODE(False)
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

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

pcDemographics.Visible = False
pcStatusDates.Visible = False
pcDependents.Visible = False
pcBanking.Visible = False
pcOtherInfo.Visible = False
pcFlags.Visible = False
pcPositionHist.Visible = False
pcSalaryHist.Visible = False
pcPerformanceHist.Visible = False
pcAttendance.Visible = False
pcAssociations.Visible = False
pcContEdu.Visible = False
pcUserDefined.Visible = False
pcFollowUps.Visible = False
pcCounseling.Visible = False
pcComments.Visible = False
pcJobMaster.Visible = False
pcPositionMaster.Visible = False
pcDashboard.Visible = False     'Ticket #22825
'Release 8.0 - Ticket #22682: Add to Label Master
pcProvince.Visible = False
'Ticket #25015 - Macaulay: New Additional Payroll ID Data
pcAddPayrollIDData.Left = 0

If Me.Caption = "Label - Demographics screen" Then
    pcDemographics.Visible = True
    pnDemographics.Visible = True
ElseIf Me.Caption = "Label - Status/Dates screen" Then
    pcStatusDates.Visible = True
    pnStatusDates.Visible = True
ElseIf Me.Caption = "Label - Dependents screen" Then
    pcDependents.Visible = True
    pnDependents.Visible = True
ElseIf Me.Caption = "Label - Banking Information screen" Then
    pcBanking.Visible = True
    pnBanking.Visible = True
ElseIf Me.Caption = "Label - Other Information screen" Then
    pcOtherInfo.Visible = True
    pnOtherInfo.Visible = True
ElseIf Me.Caption = "Setup Employee Flags" Or Me.Caption = "Label - Employee Flags screen" Then
    pcFlags.Visible = True
    frmFlags.Visible = True
ElseIf Me.Caption = "Label - Position screen" Then
    pcPositionHist.Visible = True
    pnPositionHist.Visible = True
ElseIf Me.Caption = "Label - Salary screen" Then
    pcSalaryHist.Visible = True
    pnSalaryHist.Visible = True
ElseIf Me.Caption = "Label - Performance screen" Then
    pcPerformanceHist.Visible = True
    pnPerformanceHist.Visible = True
ElseIf Me.Caption = "Label - Attendance screen" Then
    pcAttendance.Visible = True
    pnAttendance.Visible = True
ElseIf Me.Caption = lStr("Label - Associations screen") Then
    pcAssociations.Visible = True
    pnAssociations.Visible = True
ElseIf Me.Caption = "Label - Continuing Education screen" Then
    pcContEdu.Visible = True
    frmGeneral(2).Visible = True
ElseIf Me.Caption = lStr("Label - User Defined Table screen") Then
    pcUserDefined.Visible = True
    frmGeneral(3).Visible = True
ElseIf Me.Caption = lStr("Label - Follow-ups screen") Then
    pcFollowUps.Visible = True
    pnFollowUps.Visible = True
ElseIf Me.Caption = lStr("Label - Counseling screen") Then
    pcCounseling.Visible = True
    pnCounseling.Visible = True
ElseIf Me.Caption = lStr("Label - Comments screen") Then
    pcComments.Visible = True
    pnComments.Visible = True
ElseIf Me.Caption = "Label - Job Master screen" Then 'Ticket #26254 Franks 12/09/2014
    pcJobMaster.Visible = True
    pnJobMaster.Visible = True
ElseIf Me.Caption = "Label - Position Master screen" Then
    pcPositionMaster.Visible = True
    pnPositionMaster.Visible = True
ElseIf Me.Caption = "Label - Dashboard Setup screen" Then   'Ticket #22825
    pcDashboard.Visible = True
    pnDashboard.Visible = True
ElseIf Me.Caption = lStr("Label - Province/State Master screen") Then    'Release 8.0 - Ticket #22682: Add to Label Master
    pcProvince.Visible = True
    pnProvince.Visible = True
ElseIf Me.Caption = lStr("Label - Additional Payroll ID Data screen") Then   'Ticket #25015 - Macaulay: New Additional Payroll ID Data
    pcAddPayrollIDData.Visible = True
    pnAddPayrollIDData.Visible = True
End If

For X = 1 To txtNew.count
    txtNew(X).Enabled = TF
Next
'FrmDetails.Enabled = TF
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
    If Index = 0 Then
        Dim X
        glbLabLang = clpCode(0).Text
        Call setLabels
        X = EERetrieve
    End If
End Sub

Private Sub cmdPageLeft_Click(Index As Integer)
Call cmdOK_Click
Call panDisp(Index - 1)
End Sub

Private Sub cmdPageRight_Click(Index As Integer)
Call cmdOK_Click
Call panDisp(Index + 1)
End Sub

Private Sub cmdSave_Click()
    Dim rsLabel As New ADODB.Recordset
    Dim SQLQ As String
    Dim bmk As Variant
    
    'Ticket #22825
    If data2.Recordset.EOF And data2.Recordset.BOF Then
        bmk = 0
    Else
        bmk = data2.Recordset.Bookmark
    End If
    
    SQLQ = "SELECT * FROM HRLABEL WHERE LB_ORG = '" & txtItemCode.Text & "' AND LB_LANG = '" & clpCode(0).Text & "'"
    rsLabel.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsLabel.EOF Then
        'Add new dashboard item label
        rsLabel.AddNew
        rsLabel("LB_COMPNO") = "001"
        rsLabel("LB_LANG_TABL") = "LBLG"
        rsLabel("LB_LANG") = clpCode(0).Text
        rsLabel("LB_ORG") = txtItemCode.Text
    End If
    'Update the rest of dashboard item label
    rsLabel("LB_NEW") = txtNewItem.Text
    rsLabel("LB_LDATE") = Date
    rsLabel("LB_LTIME") = Time$
    rsLabel("LB_LUSER") = glbUserID
    rsLabel.Update
    
    ONewItem = txtNewItem.Text
    
    rsLabel.Close
    Set rsLabel = Nothing
    
    'Refresh Data1
    data2.Refresh
    If Not bmk = 0 Then
        data2.Recordset.Bookmark = bmk
    End If

End Sub

Private Sub cmdSave_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdUndo_Click()
    Dim bmk As Variant

    'Ticket #22825
    If data2.Recordset.EOF And data2.Recordset.BOF Then
        bmk = 0
    Else
        bmk = data2.Recordset.Bookmark
    End If

    'Refresh Data2
    data2.Refresh
    If Not bmk = 0 Then
        data2.Recordset.Bookmark = bmk
    End If
    
    'Store original values
    ONewItem = txtNewItem.Text
    
End Sub

Private Sub cmdUndo_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
glbOnTop = Me.name
fglbNew = False

Call SET_UP_MODE

Me.cmdModify_Click

End Sub

Private Sub Form_Load()
Dim X

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = Me.name

Data1.ConnectionString = glbAdoIHRDB

X = EERetrieve

For X = 1 To txtNew.count
    txtNew(X).Tag = "00-Type new label for " & txtNew(X).DataField
Next

fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

If Not gSec_Upd_Label Then
'    cmdModify.Enabled = False
End If


'ticket# 17471
frmFlags.Left = 0
frmFlags.Top = 200  '600
frmFlags.Width = 10800
frmFlags.Height = 8095

pcFlags.Left = 0
pcFlags.Top = 600
pcFlags.Width = 10800
pcFlags.Height = 8095

frmGeneral(2).Left = 0
frmGeneral(2).Top = 200   '600
frmGeneral(2).Width = 10950
frmGeneral(2).Height = 10000

pcContEdu.Left = 0
pcContEdu.Top = 600
pcContEdu.Width = 10950
pcContEdu.Height = 10000 '4815

frmGeneral(3).Left = 0
frmGeneral(3).Top = 200 '600
frmGeneral(3).Width = 11550
frmGeneral(3).Height = 8095

pcUserDefined.Left = 0
pcUserDefined.Top = 600
pcUserDefined.Width = 11550
pcUserDefined.Height = 8095

pnAssociations.Left = 0
pnAssociations.Top = 200  '600
pnAssociations.Width = 10800
pnAssociations.Height = 8095

pcAssociations.Left = 0
pcAssociations.Top = 600
pcAssociations.Width = 10800
pcAssociations.Height = 8095

pnAttendance.Left = 0
pnAttendance.Top = 200    '600
pnAttendance.Width = 10800
pnAttendance.Height = 8095

pcAttendance.Left = 0
pcAttendance.Top = 600
pcAttendance.Width = 10800
pcAttendance.Height = 8095

pnBanking.Left = 0
pnBanking.Top = 200   '600
pnBanking.Width = 10800
pnBanking.Height = 8095

pcBanking.Left = 0
pcBanking.Top = 600
pcBanking.Width = 10800
pcBanking.Height = 8095

pnComments.Left = 0
pnComments.Top = 200  '600
pnComments.Width = 10800
pnComments.Height = 8095

pcComments.Left = 0
pcComments.Top = 600
pcComments.Width = 10800
pcComments.Height = 8095

pnCounseling.Left = 0
pnCounseling.Top = 200    '600
pnCounseling.Width = 10800
pnCounseling.Height = 8095

pcCounseling.Left = 0
pcCounseling.Top = 600
pcCounseling.Width = 10800
pcCounseling.Height = 8095

pnDemographics.Left = 0
pnDemographics.Top = 200  '600
pnDemographics.Width = 10800
pnDemographics.Height = 8095

pcDemographics.Left = 0
pcDemographics.Top = 600
pcDemographics.Width = 10800
pcDemographics.Height = 8095

pnDependents.Left = 0
pnDependents.Top = 200    '600
pnDependents.Width = 10800
pnDependents.Height = 8095

pcDependents.Left = 0
pcDependents.Top = 600
pcDependents.Width = 10800
pcDependents.Height = 8095

pnFollowUps.Left = 0
pnFollowUps.Top = 200     '600
pnFollowUps.Width = 10800
pnFollowUps.Height = 8095

pcFollowUps.Left = 0
pcFollowUps.Top = 600
pcFollowUps.Width = 10800
pcFollowUps.Height = 8095

pnOtherInfo.Left = 0
pnOtherInfo.Top = 200   '600
pnOtherInfo.Width = 10800
pnOtherInfo.Height = 8095

pcOtherInfo.Left = 0
pcOtherInfo.Top = 600
pcOtherInfo.Width = 10800
pcOtherInfo.Height = 8095

pnPerformanceHist.Left = 0
pnPerformanceHist.Top = 200   '600
pnPerformanceHist.Width = 10800
pnPerformanceHist.Height = 8095

pcPerformanceHist.Left = 0
pcPerformanceHist.Top = 600
pcPerformanceHist.Width = 10800
pcPerformanceHist.Height = 8095

pnPositionHist.Left = 0
pnPositionHist.Top = 100  '600
pnPositionHist.Width = 10800
pnPositionHist.Height = 8095

pcPositionHist.Left = 0
pcPositionHist.Top = 600
pcPositionHist.Width = 10800
pcPositionHist.Height = 8095

'Ticket #26254 Franks 12/09/2014 - begin
pnJobMaster.Left = 0
pnJobMaster.Top = 200    '600
pnJobMaster.Width = 10800
pnJobMaster.Height = 8095 '3015

pcJobMaster.Left = 0
pcJobMaster.Top = 600
pcJobMaster.Width = 10800
pcJobMaster.Height = 8095
'Ticket #26254 Franks 12/09/2014 - end

pnPositionMaster.Left = 0
pnPositionMaster.Top = 200    '600
pnPositionMaster.Width = 10800
pnPositionMaster.Height = 8095

pcPositionMaster.Left = 0
pcPositionMaster.Top = 600
pcPositionMaster.Width = 10800
pcPositionMaster.Height = 8095

pnSalaryHist.Left = 0
pnSalaryHist.Top = 200    '600
pnSalaryHist.Width = 10800
pnSalaryHist.Height = 8095

pcSalaryHist.Left = 0
pcSalaryHist.Top = 600
pcSalaryHist.Width = 10800
pcSalaryHist.Height = 8095

pnStatusDates.Left = 0
pnStatusDates.Top = 200   '600
pnStatusDates.Width = 10800
pnStatusDates.Height = 11620

pcStatusDates.Left = 0
pcStatusDates.Top = 600
pcStatusDates.Width = 10800
pcStatusDates.Height = 11620

'Ticket #22825
pnDashboard.Left = 0
pnDashboard.Top = 600
pnDashboard.Width = 12335
pnDashboard.Height = 8095

pcDashboard.Left = 0
pcDashboard.Top = 600
pcDashboard.Width = 12335
pcDashboard.Height = 8095

'Release 8.0 - Ticket #22682: Add to Label Master
pnProvince.Left = 0
pnProvince.Top = 200  '600
pnProvince.Width = 10800
pnProvince.Height = 8095

pcProvince.Left = 0
pcProvince.Top = 600
pcProvince.Width = 10800
pcProvince.Height = 8095

'Ticket #25015 - Macaulay: New Additional Payroll ID Data
pnAddPayrollIDData.Left = 0
pnAddPayrollIDData.Top = 200     '600
pnAddPayrollIDData.Width = 10800
pnAddPayrollIDData.Height = 8095

pcAddPayrollIDData.Left = 0
pcAddPayrollIDData.Top = 600
pcAddPayrollIDData.Width = 10800
pcAddPayrollIDData.Height = 8095

'Jerry wants this for everyone.
'Ticket #24164 - Re-ordering and new Organizaton fields
'Samuel only
'If (glbCompSerial = "S/N - 2382W") Then
    lblOrg(185).Visible = True
    lblOrg(186).Visible = True
    txtNew(185).Visible = True
    txtNew(186).Visible = True
    lblArrow(185).Visible = True
    lblArrow(186).Visible = True
'End If


If glbWFC Then 'Ticket #25911 Franks 10/20/2014
    lblOrg(184).Visible = True: lblArrow(199).Visible = True: txtNew(199).Visible = True
    lblOrg(187).Visible = True: lblArrow(200).Visible = True: txtNew(200).Visible = True
    
    lblOrg(188).Visible = True: lblArrow(201).Visible = True: txtNew(201).Visible = True
    lblOrg(189).Visible = True: lblArrow(202).Visible = True: txtNew(202).Visible = True
End If

Call INI_Controls(Me)

End Sub

Private Function EERetrieve()
On Error Resume Next

If glbLabLang = "" Then glbLabLang = "EN"

Data1.RecordSource = "SELECT * FROM HRLABEL WHERE LB_LANG = '" & glbLabLang & "'"
Data1.Refresh
'If Not data1.Recordset.EOF Then
    For X = 1 To glbLabels(1).count
        txtNew(X).Text = glbLabels(2)(glbLabels(1)(X))
    Next
    For X = 1 To txtNew.count
        'Not the Employee Flags
        If txtNew(X).Text = "" And (X < 39 Or X > 58) Then
            txtNew(X).Text = txtNew(X).DataField
        End If
    Next
    clpCode(0).Text = Data1.Recordset("LB_LANG")
    If Len(clpCode(0).Text) = 0 Then clpCode(0).Text = "EN"
    
    'Ticket #22825- Retrieve Dashboard Items records on the Grid.
    Call Load_Dashboard_Items
'End If
End Function

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
    'panWindow.Height = Me.ScaleHeight - (panEEDESC.Height + SSPanel1.Height + panControls.Height + 200)
    'panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
    If Me.Height >= 12060 Then
        scrControl.Value = 0
        
        pcAssociations.Top = 600
        pcAttendance.Top = 600
        pcBanking.Top = 600
        pcComments.Top = 600
        pcCounseling.Top = 600
        pcDemographics.Top = 600
        pcDependents.Top = 600
        pcFollowUps.Top = 600
        pcOtherInfo.Top = 600
        pcPerformanceHist.Top = 600
        pcPositionHist.Top = 600
        pcJobMaster.Top = 600
        pcPositionMaster.Top = 600
        pcSalaryHist.Top = 600
        pcStatusDates.Top = 600
        pcContEdu.Top = 600
        pcUserDefined.Top = 600
        pcFlags.Top = 600
        pcDashboard.Top = 600
        'Release 8.0 - Ticket #22682: Add to Label Master
        pcProvince.Top = 600
        'Ticket #25015 - Macaulay: New Additional Payroll ID Data
        pcAddPayrollIDData.Top = 600
      
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 950
        
        If pcStatusDates.Visible = True Then
            scrControl.Max = 4000
        Else
            scrControl.Max = 2000
        End If
        
    End If


    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 200
    If Me.Width >= 11190 Then '9700 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        scrHScroll.Top = Me.Height - 900
        scrHScroll.Width = Me.Width - 250
    End If
    
End If

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "HRLABEL", "edit/Add")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'glbLabLang = "EN"
End Sub

Private Sub scrControl_Change()
    'panDetails(C).Top = 0 - scrControl.Value
    
    pcAssociations.Top = 600 - scrControl.Value
    pcAttendance.Top = 600 - scrControl.Value
    pcBanking.Top = 600 - scrControl.Value
    pcComments.Top = 600 - scrControl.Value
    pcCounseling.Top = 600 - scrControl.Value
    pcDemographics.Top = 600 - scrControl.Value
    pcDependents.Top = 600 - scrControl.Value
    pcFollowUps.Top = 600 - scrControl.Value
    pcOtherInfo.Top = 600 - scrControl.Value
    pcPerformanceHist.Top = 600 - scrControl.Value
    pcPositionHist.Top = 600 - scrControl.Value
    pcJobMaster.Top = 600 - scrControl.Value
    pcPositionMaster.Top = 600 - scrControl.Value
    pcSalaryHist.Top = 600 - scrControl.Value
    pcStatusDates.Top = 600 - scrControl.Value
    pcContEdu.Top = 600 - scrControl.Value
    pcUserDefined.Top = 600 - scrControl.Value
    pcFlags.Top = 600 - scrControl.Value
    pcDashboard.Top = 600 - scrControl.Value    'Ticket #22825
    'Release 8.0 - Ticket #22682: Add to Label Master
    pcProvince.Top = 600 - scrControl.Value
    'Ticket #25015 - Macaulay: New Additional Payroll ID Data
    pcAddPayrollIDData.Top = 600 - scrControl.Value
    
    
End Sub

Private Sub scrHScroll_Change()

    'panDetails(C).Left = 0 - (scrHScroll.Value / 80) * ScaleWidth
    
    
    pcAssociations.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcAttendance.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcBanking.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcComments.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcCounseling.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcDemographics.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcDependents.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcFollowUps.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcOtherInfo.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcPerformanceHist.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcPositionHist.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcJobMaster.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcPositionMaster.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcSalaryHist.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcStatusDates.Left = 0 - (scrHScroll.Value / 100) * Me.ScaleWidth
    pcContEdu.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcUserDefined.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcFlags.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    pcDashboard.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth    'Ticket #22825
    'Release 8.0 - Ticket #22682: Add to Label Master
    pcProvince.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
    'Ticket #25015 - Macaulay: New Additional Payroll ID Data
    pcAddPayrollIDData.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtNew_GotFocus(Index As Integer)
 Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtNew_KeyPress(Index As Integer, KeyAscii As Integer)
    'If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = """" Then KeyAscii = 0
End Sub

Private Sub txtNew_LostFocus(Index As Integer)
If Trim(txtNew(Index).Text) = "" And (Index < 39 Or Index > 58) Then
    txtNew(Index) = txtNew(Index).DataField
End If
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
UpdateRight = gSec_Upd_Label
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
If fglbNew Then
    UpdateState = NewRecord
    TF = True
'ElseIf data1.Recordset.EOF Then
'    UpdateState = NoRecord
'    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Public Function isChangedLabel()
On Error GoTo err_isChangedLabel
isChangedLabel = False
For X = 1 To glbLabels(1).count
    If X < 39 Or X > 58 Then
        If txtNew(X).Text <> glbLabels(2)(glbLabels(1)(X)) Then
            isChangedLabel = True
            Exit Function
        End If
    Else
        'Employee Flags
        If Len(txtNew(X).Text) > 0 Then
            If txtNew(X).Text <> glbLabels(2)(glbLabels(1)(X)) Then 'txtNew(x).DataField Then
                isChangedLabel = True
                Exit Function
            End If
        End If
    End If
Next

'Ticket #22825
'If it's Dashboard Item - then check for values changed & not saved differently
If Me.Caption = "Label - Dashboard Setup screen" Then
    If ONewItem <> txtNewItem.Text Then
        Response% = MsgBox("Do you want to Save changes?", MB_YESNO, "Save Changes?")    ' Get user response.
        If Response% = IDYES Then     ' Evaluate response
            'Save the changes
            Call cmdSave_Click
        End If
    End If
End If

Exit Function
err_isChangedLabel:
    If Err.Number = 5 Then
        If txtNew(X).Text <> txtNew(X).DataField Then
            isChangedLabel = True
            Exit Function
        End If
    End If
    Exit Function
End Function

Private Sub panDisp(Index As Integer)
Dim X%, WExit%

For X% = 0 To 3 '1
    frmGeneral(X%).Visible = False
Next
frmGeneral(Index).Visible = True

End Sub

Private Sub Load_Dashboard_Items()
    Dim rsDashboard As New ADODB.Recordset
    Dim SQLQ As String
    
    'Ticket #22825
    data2.ConnectionString = glbAdoIHRDB
    data2.RecordSource = "SELECT HR_DASHBOARD_RULE.*, HRLABEL.* FROM (HR_DASHBOARD_RULE LEFT JOIN HRLABEL ON HR_DASHBOARD_RULE.DB_ITEM_CODE = HRLABEL.LB_ORG AND LB_LANG  = '" & clpCode(0).Text & "') ORDER BY HR_DASHBOARD_RULE.DB_USERID, HR_DASHBOARD_RULE.DB_CATEGORY"
    data2.Refresh

'    SQLQ = "SELECT * FROM HR_DASHBOARD_RULE"
'    SQLQ = SQLQ & " WHERE DB_ID = " & data1.Recordset!DB_ID
'    If rsDashboard.State <> 0 Then: If rsDashboard.EOF Then rsDashboard.Close Else If rsDashboard.EditMode = adEditAdd Then rsDashboard.CancelUpdate: rsDashboard.Close Else rsDashboard.Close
'    rsDashboard.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'
'    data1.RecordSource = SQLQ
'    data1.Refresh

    'Set the Original Label to New Item if New Item is blank
    If Not data2.Recordset.EOF Then
        If IsNull(data2.Recordset("LB_NEW")) Or data2.Recordset("LB_NEW") = "" Then
            txtNewItem.Text = data2.Recordset("DB_ITEM_DESC")
        Else
            txtNewItem.Text = data2.Recordset("LB_NEW")
        End If
        
        'Retain the original value to see later if it changed.
        ONewItem = txtNewItem.Text
    End If
    
    vbxTrueGrid.Columns(0).Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(5).Visible = False
    vbxTrueGrid.Columns(6).Visible = False
End Sub

Private Sub txtNewItem_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
    Dim Response%

    If ONewItem <> txtNewItem.Text Then
        Response% = MsgBox("Do you want to Save changes?", MB_YESNO, "Save Changes?")    ' Get user response.
        If Response% = IDYES Then     ' Evaluate response
            'Save the changes
            Call cmdSave_Click
        End If
    Else
        Cancel = False
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    'Ticket #22825
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT HR_DASHBOARD_RULE.*, HRLABEL.* FROM (HR_DASHBOARD_RULE LEFT JOIN HRLABEL ON HR_DASHBOARD_RULE.DB_ITEM_CODE = HRLABEL.LB_ORG AND LB_LANG  = '" & clpCode(0).Text & "')"
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    data2.RecordSource = SQLQ
    data2.Refresh

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Ticket #22825
    'Set the Original Label to New Item if New Item is blank
    If Not data2.Recordset.EOF Then
        lblArrow(175).Left = lblItemDesc.Left + lblItemDesc.Width + 380
        txtNewItem.Left = lblArrow(175).Left + 480
        If IsNull(data2.Recordset("LB_NEW")) Or data2.Recordset("LB_NEW") = "" Then
            txtNewItem.Text = data2.Recordset("DB_ITEM_DESC")
        Else
            txtNewItem.Text = data2.Recordset("LB_NEW")
        End If
        
        'Retain the original value to see later if it changed.
        ONewItem = txtNewItem.Text
    End If
        
End Sub
