VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEmployeeFlags 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Flags"
   ClientHeight    =   11010
   ClientLeft      =   135
   ClientTop       =   2535
   ClientWidth     =   14280
   ForeColor       =   &H00C0C0C0&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   32562.85
   ScaleMode       =   0  'User
   ScaleWidth      =   54827.9
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHorz 
      Height          =   321
      LargeChange     =   315
      Left            =   0
      Max             =   14415
      SmallChange     =   315
      TabIndex        =   157
      Top             =   9720
      Width           =   15495
   End
   Begin Threed.SSPanel panWindow 
      Height          =   9015
      Left            =   120
      TabIndex        =   110
      Top             =   600
      Width           =   15855
      _Version        =   65536
      _ExtentX        =   27966
      _ExtentY        =   15901
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
      Begin VB.PictureBox panDetails 
         BorderStyle     =   0  'None
         Height          =   8775
         Left            =   240
         ScaleHeight     =   8775
         ScaleWidth      =   15735
         TabIndex        =   111
         Top             =   120
         Width           =   15735
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   19
            Left            =   14880
            TabIndex        =   99
            Tag             =   "Click to select the location"
            Top             =   7917
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   18
            Left            =   14880
            TabIndex        =   94
            Tag             =   "Click to select the location"
            Top             =   7512
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   17
            Left            =   14880
            TabIndex        =   89
            Tag             =   "Click to select the location"
            Top             =   7122
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   16
            Left            =   14880
            TabIndex        =   84
            Tag             =   "Click to select the location"
            Top             =   6717
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   15
            Left            =   14880
            TabIndex        =   79
            Tag             =   "Click to select the location"
            Top             =   6327
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   14
            Left            =   14880
            TabIndex        =   74
            Tag             =   "Click to select the location"
            Top             =   5937
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   13
            Left            =   14880
            TabIndex        =   69
            Tag             =   "Click to select the location"
            Top             =   5532
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   12
            Left            =   14880
            TabIndex        =   64
            Tag             =   "Click to select the location"
            Top             =   5142
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   11
            Left            =   14880
            TabIndex        =   59
            Tag             =   "Click to select the location"
            Top             =   4737
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   10
            Left            =   14880
            TabIndex        =   54
            Tag             =   "Click to select the location"
            Top             =   4347
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   9
            Left            =   14880
            TabIndex        =   49
            Tag             =   "Click to select the location"
            Top             =   3942
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   8
            Left            =   14880
            TabIndex        =   44
            Tag             =   "Click to select the location"
            Top             =   3552
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   7
            Left            =   14880
            TabIndex        =   39
            Tag             =   "Click to select the location"
            Top             =   3147
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   6
            Left            =   14880
            TabIndex        =   34
            Tag             =   "Click to select the location"
            Top             =   2757
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   5
            Left            =   14880
            TabIndex        =   29
            Tag             =   "Click to select the location"
            Top             =   2352
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   4
            Left            =   14880
            TabIndex        =   24
            Tag             =   "Click to select the location"
            Top             =   1962
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   3
            Left            =   14880
            TabIndex        =   19
            Tag             =   "Click to select the location"
            Top             =   1557
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   2
            Left            =   14880
            TabIndex        =   14
            Tag             =   "Click to select the location"
            Top             =   1167
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   1
            Left            =   14880
            TabIndex        =   9
            Tag             =   "Click to select the location"
            Top             =   762
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   290
            Index           =   0
            Left            =   14880
            TabIndex        =   4
            Tag             =   "Click to select the location"
            Top             =   372
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   0
            Left            =   3120
            TabIndex        =   0
            Tag             =   "00-Enter Value for this Flag"
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL20"
            Height          =   285
            Index           =   19
            Left            =   6000
            TabIndex        =   131
            Top             =   7920
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL19"
            Height          =   285
            Index           =   18
            Left            =   6000
            TabIndex        =   130
            Top             =   7530
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL18"
            Height          =   285
            Index           =   17
            Left            =   6000
            TabIndex        =   129
            Top             =   7125
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL17"
            Height          =   285
            Index           =   16
            Left            =   6000
            TabIndex        =   128
            Top             =   6735
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL16"
            Height          =   285
            Index           =   15
            Left            =   6000
            TabIndex        =   127
            Top             =   6330
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL15"
            Height          =   285
            Index           =   14
            Left            =   6000
            TabIndex        =   126
            Top             =   5925
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL14"
            Height          =   285
            Index           =   13
            Left            =   6000
            TabIndex        =   125
            Top             =   5535
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL13"
            Height          =   285
            Index           =   12
            Left            =   6000
            TabIndex        =   124
            Top             =   5130
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL12"
            Height          =   285
            Index           =   11
            Left            =   6000
            TabIndex        =   123
            Top             =   4725
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL11"
            Height          =   285
            Index           =   10
            Left            =   6000
            TabIndex        =   122
            Top             =   4305
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL10"
            Height          =   285
            Index           =   9
            Left            =   6000
            TabIndex        =   121
            Top             =   3960
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL9"
            Height          =   285
            Index           =   8
            Left            =   6000
            TabIndex        =   120
            Top             =   3570
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL8"
            Height          =   285
            Index           =   7
            Left            =   6000
            TabIndex        =   119
            Top             =   3165
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL7"
            Height          =   285
            Index           =   6
            Left            =   6000
            TabIndex        =   118
            Top             =   2775
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL6"
            Height          =   285
            Index           =   5
            Left            =   6000
            TabIndex        =   117
            Top             =   2370
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL5"
            Height          =   285
            Index           =   4
            Left            =   6000
            TabIndex        =   116
            Top             =   1950
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL4"
            Height          =   285
            Index           =   3
            Left            =   6000
            TabIndex        =   115
            Top             =   1575
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL3"
            Height          =   285
            Index           =   2
            Left            =   6000
            TabIndex        =   114
            Top             =   1170
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL2"
            Height          =   285
            Index           =   1
            Left            =   6000
            TabIndex        =   113
            Top             =   765
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtEmpFlag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            DataField       =   "EF_FLAGVAL1"
            Height          =   285
            Index           =   0
            Left            =   6000
            TabIndex        =   112
            Top             =   360
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   19
            Left            =   3120
            TabIndex        =   95
            Tag             =   "00-Enter Value for this Flag"
            Top             =   7905
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   18
            Left            =   3120
            TabIndex        =   90
            Tag             =   "00-Enter Value for this Flag"
            Top             =   7500
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   17
            Left            =   3120
            TabIndex        =   85
            Tag             =   "00-Enter Value for this Flag"
            Top             =   7110
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   16
            Left            =   3120
            TabIndex        =   80
            Tag             =   "00-Enter Value for this Flag"
            Top             =   6705
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   15
            Left            =   3120
            TabIndex        =   75
            Tag             =   "00-Enter Value for this Flag"
            Top             =   6315
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   14
            Left            =   3120
            TabIndex        =   70
            Tag             =   "00-Enter Value for this Flag"
            Top             =   5925
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   13
            Left            =   3120
            TabIndex        =   65
            Tag             =   "00-Enter Value for this Flag"
            Top             =   5520
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   12
            Left            =   3120
            TabIndex        =   60
            Tag             =   "00-Enter Value for this Flag"
            Top             =   5130
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   11
            Left            =   3120
            TabIndex        =   55
            Tag             =   "00-Enter Value for this Flag"
            Top             =   4725
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   10
            Left            =   3120
            TabIndex        =   50
            Tag             =   "00-Enter Value for this Flag"
            Top             =   4335
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   9
            Left            =   3120
            TabIndex        =   45
            Tag             =   "00-Enter Value for this Flag"
            Top             =   3930
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   8
            Left            =   3120
            TabIndex        =   40
            Tag             =   "00-Enter Value for this Flag"
            Top             =   3540
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   7
            Left            =   3120
            TabIndex        =   35
            Tag             =   "00-Enter Value for this Flag"
            Top             =   3135
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   6
            Left            =   3120
            TabIndex        =   30
            Tag             =   "00-Enter Value for this Flag"
            Top             =   2745
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   5
            Left            =   3120
            TabIndex        =   25
            Tag             =   "00-Enter Value for this Flag"
            Top             =   2340
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   15
            Tag             =   "00-Enter Value for this Flag"
            Top             =   1545
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   2
            Left            =   3120
            TabIndex        =   10
            Tag             =   "00-Enter Value for this Flag"
            Top             =   1155
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   20
            Tag             =   "00-Enter Value for this Flag"
            Top             =   1950
            Width           =   2775
         End
         Begin VB.ComboBox cboEmpFlag 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   5
            Tag             =   "00-Enter Value for this Flag"
            Top             =   750
            Width           =   2775
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE1"
            Height          =   315
            Index           =   0
            Left            =   6480
            TabIndex        =   1
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE2"
            Height          =   315
            Index           =   1
            Left            =   6480
            TabIndex        =   6
            Top             =   750
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE3"
            Height          =   315
            Index           =   2
            Left            =   6480
            TabIndex        =   11
            Top             =   1155
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE4"
            Height          =   315
            Index           =   3
            Left            =   6480
            TabIndex        =   16
            Top             =   1545
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE5"
            Height          =   315
            Index           =   4
            Left            =   6480
            TabIndex        =   21
            Top             =   1950
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE6"
            Height          =   315
            Index           =   5
            Left            =   6480
            TabIndex        =   26
            Top             =   2340
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE7"
            Height          =   315
            Index           =   6
            Left            =   6480
            TabIndex        =   31
            Top             =   2745
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE8"
            Height          =   315
            Index           =   7
            Left            =   6480
            TabIndex        =   36
            Top             =   3135
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE9"
            Height          =   315
            Index           =   8
            Left            =   6480
            TabIndex        =   41
            Top             =   3540
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE10"
            Height          =   315
            Index           =   9
            Left            =   6480
            TabIndex        =   46
            Top             =   3930
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE11"
            Height          =   315
            Index           =   10
            Left            =   6480
            TabIndex        =   51
            Top             =   4335
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE12"
            Height          =   315
            Index           =   11
            Left            =   6480
            TabIndex        =   56
            Top             =   4725
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE13"
            Height          =   315
            Index           =   12
            Left            =   6480
            TabIndex        =   61
            Top             =   5130
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE14"
            Height          =   315
            Index           =   13
            Left            =   6480
            TabIndex        =   66
            Top             =   5520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE15"
            Height          =   315
            Index           =   14
            Left            =   6480
            TabIndex        =   71
            Top             =   5925
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE16"
            Height          =   315
            Index           =   15
            Left            =   6480
            TabIndex        =   76
            Top             =   6315
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE17"
            Height          =   315
            Index           =   16
            Left            =   6480
            TabIndex        =   81
            Top             =   6705
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE18"
            Height          =   315
            Index           =   17
            Left            =   6480
            TabIndex        =   86
            Top             =   7110
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE19"
            Height          =   315
            Index           =   18
            Left            =   6480
            TabIndex        =   91
            Top             =   7500
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpDate 
            DataField       =   "EF_FLAGDTE20"
            Height          =   315
            Index           =   19
            Left            =   6480
            TabIndex        =   96
            Top             =   7905
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE1"
            Height          =   315
            Index           =   0
            Left            =   8400
            TabIndex        =   2
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE2"
            Height          =   315
            Index           =   1
            Left            =   8400
            TabIndex        =   7
            Top             =   750
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE3"
            Height          =   315
            Index           =   2
            Left            =   8400
            TabIndex        =   12
            Top             =   1155
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE4"
            Height          =   315
            Index           =   3
            Left            =   8400
            TabIndex        =   17
            Top             =   1545
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE5"
            Height          =   315
            Index           =   4
            Left            =   8400
            TabIndex        =   22
            Top             =   1950
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE6"
            Height          =   315
            Index           =   5
            Left            =   8400
            TabIndex        =   27
            Top             =   2340
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE7"
            Height          =   315
            Index           =   6
            Left            =   8400
            TabIndex        =   32
            Top             =   2745
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE8"
            Height          =   315
            Index           =   7
            Left            =   8400
            TabIndex        =   37
            Top             =   3135
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE9"
            Height          =   315
            Index           =   8
            Left            =   8400
            TabIndex        =   42
            Top             =   3540
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE10"
            Height          =   315
            Index           =   9
            Left            =   8400
            TabIndex        =   47
            Top             =   3930
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE11"
            Height          =   315
            Index           =   10
            Left            =   8400
            TabIndex        =   52
            Top             =   4335
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE12"
            Height          =   315
            Index           =   11
            Left            =   8400
            TabIndex        =   57
            Top             =   4725
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE13"
            Height          =   315
            Index           =   12
            Left            =   8400
            TabIndex        =   62
            Top             =   5130
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE14"
            Height          =   315
            Index           =   13
            Left            =   8400
            TabIndex        =   67
            Top             =   5520
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE15"
            Height          =   315
            Index           =   14
            Left            =   8400
            TabIndex        =   72
            Top             =   5925
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE16"
            Height          =   315
            Index           =   15
            Left            =   8400
            TabIndex        =   77
            Top             =   6315
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE17"
            Height          =   315
            Index           =   16
            Left            =   8400
            TabIndex        =   82
            Top             =   6705
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE18"
            Height          =   315
            Index           =   17
            Left            =   8400
            TabIndex        =   87
            Top             =   7110
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE19"
            Height          =   315
            Index           =   18
            Left            =   8400
            TabIndex        =   92
            Top             =   7500
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.DateLookup dlpFU 
            DataField       =   "EF_FUDTE20"
            Height          =   315
            Index           =   19
            Left            =   8400
            TabIndex        =   97
            Top             =   7905
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            ShowDescription =   0   'False
            TextBoxWidth    =   1375
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS1"
            Height          =   285
            Index           =   0
            Left            =   10200
            TabIndex        =   3
            Tag             =   "01-Follow-up Reason"
            Top             =   375
            Visible         =   0   'False
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS2"
            Height          =   285
            Index           =   1
            Left            =   10200
            TabIndex        =   8
            Tag             =   "01-Follow-up Reason"
            Top             =   765
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS3"
            Height          =   285
            Index           =   2
            Left            =   10200
            TabIndex        =   13
            Tag             =   "01-Follow-up Reason"
            Top             =   1170
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS4"
            Height          =   285
            Index           =   3
            Left            =   10200
            TabIndex        =   18
            Tag             =   "01-Follow-up Reason"
            Top             =   1560
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS5"
            Height          =   285
            Index           =   4
            Left            =   10200
            TabIndex        =   23
            Tag             =   "01-Follow-up Reason"
            Top             =   1965
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS6"
            Height          =   285
            Index           =   5
            Left            =   10200
            TabIndex        =   28
            Tag             =   "01-Follow-up Reason"
            Top             =   2355
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS7"
            Height          =   285
            Index           =   6
            Left            =   10200
            TabIndex        =   33
            Tag             =   "01-Follow-up Reason"
            Top             =   2760
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS8"
            Height          =   285
            Index           =   7
            Left            =   10200
            TabIndex        =   38
            Tag             =   "01-Follow-up Reason"
            Top             =   3150
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS9"
            Height          =   285
            Index           =   8
            Left            =   10200
            TabIndex        =   43
            Tag             =   "01-Follow-up Reason"
            Top             =   3555
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS10"
            Height          =   285
            Index           =   9
            Left            =   10200
            TabIndex        =   48
            Tag             =   "01-Follow-up Reason"
            Top             =   3945
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS11"
            Height          =   285
            Index           =   10
            Left            =   10200
            TabIndex        =   53
            Tag             =   "01-Follow-up Reason"
            Top             =   4350
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS12"
            Height          =   285
            Index           =   11
            Left            =   10200
            TabIndex        =   58
            Tag             =   "01-Follow-up Reason"
            Top             =   4740
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS13"
            Height          =   285
            Index           =   12
            Left            =   10200
            TabIndex        =   63
            Tag             =   "01-Follow-up Reason"
            Top             =   5145
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS14"
            Height          =   285
            Index           =   13
            Left            =   10200
            TabIndex        =   68
            Tag             =   "01-Follow-up Reason"
            Top             =   5535
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS15"
            Height          =   285
            Index           =   14
            Left            =   10200
            TabIndex        =   73
            Tag             =   "01-Follow-up Reason"
            Top             =   5940
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS16"
            Height          =   285
            Index           =   15
            Left            =   10200
            TabIndex        =   78
            Tag             =   "01-Follow-up Reason"
            Top             =   6330
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS17"
            Height          =   285
            Index           =   16
            Left            =   10200
            TabIndex        =   83
            Tag             =   "01-Follow-up Reason"
            Top             =   6720
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS18"
            Height          =   285
            Index           =   17
            Left            =   10200
            TabIndex        =   88
            Tag             =   "01-Follow-up Reason"
            Top             =   7125
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FTREAS19"
            Height          =   285
            Index           =   18
            Left            =   10200
            TabIndex        =   93
            Tag             =   "01-Follow-up Reason"
            Top             =   7515
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EF_FUREAS20"
            Height          =   285
            Index           =   19
            Left            =   10200
            TabIndex        =   98
            Tag             =   "01-Follow-up Reason"
            Top             =   7920
            Visible         =   0   'False
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "FURE"
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   19
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0000
            Top             =   7942
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   19
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":014A
            Top             =   7942
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   18
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0294
            Top             =   7537
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   18
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":03DE
            Top             =   7537
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   17
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0528
            Top             =   7147
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   17
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0672
            Top             =   7147
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   16
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":07BC
            Top             =   6742
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   16
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0906
            Top             =   6742
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   15
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0A50
            Top             =   6352
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   15
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0B9A
            Top             =   6352
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   14
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0CE4
            Top             =   5962
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   14
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0E2E
            Top             =   5962
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   13
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":0F78
            Top             =   5557
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   13
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":10C2
            Top             =   5557
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   12
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":120C
            Top             =   5167
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   12
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":1356
            Top             =   5167
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   11
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":14A0
            Top             =   4762
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   11
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":15EA
            Top             =   4762
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   10
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":1734
            Top             =   4372
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   10
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":187E
            Top             =   4372
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   9
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":19C8
            Top             =   3967
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   9
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":1B12
            Top             =   3967
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   8
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":1C5C
            Top             =   3577
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   8
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":1DA6
            Top             =   3577
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   7
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":1EF0
            Top             =   3172
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   7
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":203A
            Top             =   3172
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   6
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2184
            Top             =   2782
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   6
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":22CE
            Top             =   2782
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   5
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2418
            Top             =   2377
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   5
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2562
            Top             =   2377
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   4
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":26AC
            Top             =   1987
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   4
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":27F6
            Top             =   1987
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   3
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2940
            Top             =   1582
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   3
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2A8A
            Top             =   1582
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   2
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2BD4
            Top             =   1192
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   2
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2D1E
            Top             =   1192
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   1
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2E68
            Top             =   787
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblAttach 
            AutoSize        =   -1  'True
            Caption         =   "Attachment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   14520
            TabIndex        =   159
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Index           =   0
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":2FB2
            Top             =   397
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 1"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   156
            Top             =   390
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 2"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   155
            Top             =   780
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 3"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   154
            Top             =   1185
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 4"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   153
            Top             =   1575
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 5"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   152
            Top             =   1980
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 6"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   151
            Top             =   2370
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 7"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   150
            Top             =   2775
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 8"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   149
            Top             =   3165
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 9"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   148
            Top             =   3570
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 10"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   147
            Top             =   3960
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 11"
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   146
            Top             =   4365
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 12"
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   145
            Top             =   4755
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 13"
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   144
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 14"
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   143
            Top             =   5550
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 15"
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   142
            Top             =   5955
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 16"
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   141
            Top             =   6345
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 17"
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   140
            Top             =   6735
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 18"
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   139
            Top             =   7140
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 19"
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   138
            Top             =   7530
            Width           =   3015
         End
         Begin VB.Label lblTitle 
            BackStyle       =   0  'Transparent
            Caption         =   "Employee Flag 20"
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   137
            Top             =   7935
            Width           =   3015
         End
         Begin VB.Label Label1 
            Caption         =   "Reason"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10440
            TabIndex        =   136
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Flag"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   135
            Top             =   0
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Value"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   134
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6360
            TabIndex        =   133
            Top             =   0
            Width           =   1815
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Follow-Up"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8400
            TabIndex        =   132
            Top             =   0
            Width           =   1695
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   1
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":30FC
            Top             =   787
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgSec 
            Height          =   240
            Index           =   0
            Left            =   14520
            Picture         =   "frmEmployeeFlags.frx":3246
            Top             =   397
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Width           =   14280
      _Version        =   65536
      _ExtentX        =   25188
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
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
         Left            =   8400
         TabIndex        =   158
         Top             =   120
         Width           =   1305
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
         Left            =   3150
         TabIndex        =   104
         Top             =   120
         Width           =   1740
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         DataField       =   "EF_EMPNBR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   103
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblEmpNbr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   102
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label lblEEnum 
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
         Height          =   255
         Left            =   1440
         TabIndex        =   101
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   9015
      LargeChange     =   315
      Left            =   16080
      Max             =   8775
      SmallChange     =   315
      TabIndex        =   109
      Top             =   600
      Width           =   345
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "EF_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5100
      MaxLength       =   25
      TabIndex        =   107
      TabStop         =   0   'False
      Text            =   "Ldate"
      Top             =   11790
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "EF_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   6660
      MaxLength       =   25
      TabIndex        =   106
      TabStop         =   0   'False
      Text            =   "LTime"
      Top             =   11790
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "EF_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8220
      MaxLength       =   25
      TabIndex        =   105
      TabStop         =   0   'False
      Text            =   "LUser"
      Top             =   11790
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   2100
      Top             =   11760
      Visible         =   0   'False
      Width           =   2115
      _ExtentX        =   3731
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
      Caption         =   "Data1"
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
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EF_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4320
      TabIndex        =   108
      Top             =   11880
      Visible         =   0   'False
      Width           =   690
   End
End
Attribute VB_Name = "frmEmployeeFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbNew As Integer
Dim flagFrmLoad As Boolean
Dim rsDATA As New ADODB.Recordset

Private Sub cboEmpFlag_Change(Index As Integer)
    If Trim(txtEmpFlag(Index)) <> Trim(cboEmpFlag(Index)) Then
        txtEmpFlag(Index) = Trim(cboEmpFlag(Index))
    End If
End Sub

Private Sub cboEmpFlag_LostFocus(Index As Integer)
    If Trim(txtEmpFlag(Index)) <> Trim(cboEmpFlag(Index)) Then
        txtEmpFlag(Index) = Trim(cboEmpFlag(Index))
    End If
End Sub

Private Sub clpCode_Change(Index As Integer)
'    'Call procedure to check if user has access to the follow up code
'    If Len(clpCode(Index).Text) > 0 And clpCode(Index).Caption <> "Unassigned" And clpCode(Index).Caption <> "" Then
'        If Not Accessible_FollowUp_Code(clpCode(Index).Text) Then
'            'User does not have right to access this Follow Up code - hide the Employee Flag record
'            Call Hide_Employee_Flag(Index)
'        Else
'            'User has right to access this Follow Up code - show the Employee Flag record
'            Call Show_Employee_Flag(Index)
'        End If
'    Else
'        If Len(lblTitle(Index).Caption) > 0 Then
'            'User has right to access this Follow Up code - show the Employee Flag record
'            Call Show_Employee_Flag(Index)
'        End If
'    End If
End Sub

Private Sub cmdImport_Click(Index As Integer)
    glbDocName = "EmployeeFlag"
    glbEmpFlagNo = Index
    glbEmpFlag = lblTitle(Index).Caption
    If IsDate(dlpDate(Index)) Then
        glbEmpFlagDate = dlpDate(Index)
    End If
    glbDocKey = rsDATA("EF_ID")
    frmInAttachment.Show 1
    DoEvents
    
    Call DispimgIcon(Me, "frmEmployeeFlags")
    cmdImport(Index).Visible = True
    
    If frmInAttachment.IfExist Then
        imgSec(Index).Visible = True
        imgNoSec(Index).Visible = False
    Else
        imgSec(Index).Visible = False
        imgNoSec(Index).Visible = True
    End If
End Sub

Private Sub dlpFU_Change(Index As Integer)
    If Len(dlpFU(Index).Text) > 0 Then
        clpCode(Index).Visible = True
    End If
End Sub

Private Sub Form_Activate()
    Dim X As Integer
    
    glbOnTop = "FRMEMPLOYEEFLAGS"
    Call SET_UP_MODE
    
    'Based on Follow Up Code security
    Call HideShow_EmployeeFlagsRec_FollowUp_Security
    
'    For x = 0 To 19
'        'Call procedure to check if user has access to the follow up code
'        If Len(clpCode(x).Text) > 0 And clpCode(x).Caption <> "Unassigned" And clpCode(x).Caption <> "" Then
'            If Not Accessible_FollowUp_Code(clpCode(x).Text) Then
'                'User does not have right to access this Follow Up code - hide the Employee Flag record
'                Call Hide_Employee_Flag(x)
'            Else
'                'User has right to access this Follow Up code - show the Employee Flag record
'                Call Show_Employee_Flag(x)
'            End If
'        Else
'            If Len(lblTitle(x).Caption) > 0 Then
'                'User has right to access this Follow Up code - show the Employee Flag record
'                Call Show_Employee_Flag(x)
'            End If
'        End If
'    Next
    
End Sub

Private Sub Form_Resize()
    If Not (Me.WindowState = vbMinimized Or MDIMain.WindowState = vbMinimized) Then
        panWindow.Height = Me.ScaleHeight - (panEEDESC.Height + scrHorz.Height + 200)
        panWindow.Width = Me.ScaleWidth - (scrControl.Width + 200)
        If Me.Height >= panEEDESC.Height + panDetails.Height Then   '+ 230 Then
            scrControl.Value = 0
            panDetails.Top = 0
            scrControl.Visible = False
        Else
            scrControl.Visible = True
            scrControl.Left = Me.ScaleWidth - scrControl.Width
            scrControl.Height = panWindow.Height
        End If
    
        If Me.Width >= 14415 Then   '+ 230 Then
            scrHorz.Value = 0
            panDetails.Left = 0
            scrHorz.Visible = False
        Else
            scrHorz.Visible = True
            scrHorz.Top = panWindow.Top + panWindow.Height
            scrHorz.Width = panWindow.Width
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEmployeeFlags = Nothing
    Call NextForm
End Sub

Private Sub imgSec_Click(Index As Integer)
    Dim SQLQ
    glbEmpFlagNo = Index
    If IsDate(dlpDate(Index)) Then
        glbEmpFlagDate = dlpDate(Index)
    End If
    SQLQ = getSQL("frmEmployeeFlags")
    Call FillMemoFile(SQLQ, "EmployeeFlag")
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    Me.Caption = "Employee Flags - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblEEnum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Function CheckFlags()
Dim X As Long, n As Long
Dim vList As String
    For X = 0 To lblTitle.count - 1
        lblTitle(X).Caption = lStr(lblTitle(X).Caption)
        If IsNull(lblTitle(X).Caption) Or lblTitle(X).Caption = "" Or Left(lblTitle(X).Caption, 13) = "Employee Flag" Then
            lblTitle(X).Caption = ""
            cboEmpFlag(X).Text = ""
            cboEmpFlag(X).Visible = False
            clpCode(X).Text = ""
            clpCode(X).Visible = False
            dlpDate(X).Text = ""
            dlpDate(X).Visible = False
            dlpFU(X).Text = ""
            dlpFU(X).Visible = False
        
            If gsAttachment_DB Then
                'lblAttach.Visible = False
                imgNoSec(X).Visible = False
                imgSec(X).Visible = False
                cmdImport(X).Visible = False
            End If
        
        Else
            cboEmpFlag(X).Clear
            vList = FlgList(X)
            n = 1
            Do While n > 0
                n = InStr(vList, "&")
                If n > 0 Then
                    cboEmpFlag(X).AddItem Trim(Left(vList, n - 1))
                    vList = Mid(vList, n + 1)
                Else
                    cboEmpFlag(X).AddItem Trim(vList)
                End If
            Loop
            
            'If gsAttachment_DB Then
            '    lblAttach.Visible = True
            '    imgNoSec(X).Visible = True
            '    imgSec(X).Visible = True
            '    cmdImport(X).Visible = True
            'End If
        End If
    Next
        
End Function

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer, C As Long ' records found
Dim impInd As Integer
Dim X As Integer

glbOnTop = "FRMEMPLOYEEFLAGS"
Screen.MousePointer = HOURGLASS
If glbtermopen Then
Data1.ConnectionString = glbAdoIHRAUDIT
Else
Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

Call setCaption(Label5)

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_EmpFlags_ViewOwn Then
        MsgBox "You cannot view your own Employee Flags Information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Sub
    End If
End If

Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call INI_Controls(Me)
''''''''''''''''''
Call CheckFlags

For C = 0 To cboEmpFlag.count - 1
    cboEmpFlag(C).Text = Trim(txtEmpFlag(C).Text)
Next
Data1.Refresh

fglbNew = False

Call SET_UP_MODE
''''''''''''''

'Based on Follow Up Code security
Call HideShow_EmployeeFlagsRec_FollowUp_Security

'For x = 0 To 19
'    'Call procedure to check if user has access to the follow up code
'    If Len(clpCode(x).Text) > 0 And clpCode(x).Caption <> "Unassigned" And clpCode(x).Caption <> "" Then
'        If Not Accessible_FollowUp_Code(clpCode(x).Text) Then
'            'User does not have right to access this Follow Up code - hide the Employee Flag record
'            Call Hide_Employee_Flag(x)
'        Else
'            'User has right to access this Follow Up code - show the Employee Flag record
'            Call Show_Employee_Flag(x)
'        End If
'    Else
'        If Len(lblTitle(x).Caption) > 0 Then
'            'User has right to access this Follow Up code - show the Employee Flag record
'            Call Show_Employee_Flag(x)
'        End If
'    End If
'Next

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

End Sub

Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "EF_COMPNO,EF_EMPNBR, EF_FLAG1, EF_FLAGVAL1, EF_FLAG2, "
SQLQ = SQLQ & "EF_FLAGVAL2, EF_FLAG3,EF_FLAGVAL3,EF_FLAG4,EF_FLAGVAL4,"
SQLQ = SQLQ & "EF_FLAG5, EF_FLAGVAL5, EF_FLAG6,EF_FLAGVAL6,EF_FLAG7,EF_FLAGVAL7,"
SQLQ = SQLQ & "EF_FLAG8, EF_FLAGVAL8, EF_FLAG9,EF_FLAGVAL9,EF_FLAG10,EF_FLAGVAL10,"
SQLQ = SQLQ & "EF_FLAG11, EF_FLAGVAL11, EF_FLAG12,EF_FLAGVAL12,EF_FLAG13,EF_FLAGVAL13,"
SQLQ = SQLQ & "EF_FLAG14, EF_FLAGVAL14, EF_FLAG15,EF_FLAGVAL15,EF_FLAG16,EF_FLAGVAL16,"
SQLQ = SQLQ & "EF_FLAG17, EF_FLAGVAL17, EF_FLAG18,EF_FLAGVAL18,EF_FLAG19,EF_FLAGVAL19,EF_FLAG20,EF_FLAGVAL20,"
SQLQ = SQLQ & "EF_LDATE, EF_LTIME, EF_LUSER"

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList = SQLQ
End Function

Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_EmpFlags_ViewOwn Then
        MsgBox "You cannot view your own Employee Flags Information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

If glbtermopen Then
    SQLQ = "Select * from Term_HREMP_FLAGS"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select * from HREMP_FLAGS "
    SQLQ = SQLQ & " where EF_EMPNBR = " & glbLEE_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsDATA.EOF Then
    rsDATA.AddNew
    rsDATA("EF_COMPNO") = "001"
    If glbtermopen Then
        rsDATA("EF_EMPNBR") = glbTERM_ID
        rsDATA("TERM_SEQ") = glbTERM_Seq
    Else
        rsDATA("EF_EMPNBR") = glbLEE_ID
    End If
    rsDATA.Update
End If

Data1.RecordSource = SQLQ
Data1.Refresh
Call Display_Value

'If rsDATA.BOF And rsDATA.EOF Then
'   MsgBox "Sorry, Employee Removed prior to your access"
'Else
   EERetrieve = True
'End If


Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP_FLAGS", "SELECT")
Call RollBack

Exit Function

End Function

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

Public Sub Display_Value()
    Dim SQLQ As String
    Dim X As Integer
    
    If rsDATA.EOF Or rsDATA.BOF Then
        Call Set_Control("B", Me)
        Call SET_UP_MODE
        Exit Sub
    End If
    
    If glbtermopen Then
        SQLQ = "Select * from Term_HREMP_FLAGS"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        
    Else
        SQLQ = "Select * from HREMP_FLAGS "
        SQLQ = SQLQ & " where EF_EMPNBR = " & glbLEE_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    'Data1.RecordSource = SQLQ
    'Data1.Refresh
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
    glbDocName = "EmployeeFlag"
    If gsAttachment_DB Then
        lblAttach.Visible = True
        glbDocKey = 0
        If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
            If rsDATA.RecordCount > 0 Then
                If Not IsNull(rsDATA("EF_DOCKEY")) Then
                    glbDocKey = rsDATA("EF_DOCKEY")
                Else
                    glbDocKey = 0
                End If
            Else
                If Not IsNull(Data1.Recordset("EF_DOCKEY")) Then
                    glbDocKey = Data1.Recordset("EF_DOCKEY")
                Else
                    glbDocKey = 0
                End If
            End If
        End If
        
        Call DispimgIcon(Me, "frmEmployeeFlags")
        
        For X = 0 To lblTitle.count - 1
            glbEmpFlagNo = X
            If Not IsNull(rsDATA("EF_FLAGDTE" & X + 1)) Then
                glbEmpFlagDate = rsDATA("EF_FLAGDTE" & X + 1)
            End If
            lblTitle(X).Caption = lStr(lblTitle(X).Caption)
            If Not IsNull(lblTitle(X).Caption) And lblTitle(X).Caption <> "" Then
                If frmInAttachment.IfExist Then
                    imgSec(X).Visible = True
                    imgNoSec(X).Visible = False
                Else
                    imgSec(X).Visible = False
                    imgNoSec(X).Visible = True
                End If
                
                If gSec_Upd_EMP_FLAGS And Not glbtermopen Then
                    lblAttach.Visible = True
                    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
                        cmdImport(X).Visible = False
                    Else
                        cmdImport(X).Visible = True
                    End If
                End If
            End If
        Next
    End If
    
'Based on Follow Up Code security
Call HideShow_EmployeeFlagsRec_FollowUp_Security
        
'    For x = 0 To 19
'        'Call procedure to check if user has access to the follow up code
'        If Len(clpCode(x).Text) > 0 And clpCode(x).Caption <> "Unassigned" And clpCode(x).Caption <> "" Then
'            If Not Accessible_FollowUp_Code(clpCode(x).Text) Then
'                'User does not have right to access this Follow Up code - hide the Employee Flag record
'                Call Hide_Employee_Flag(x)
'            Else
'                'User has right to access this Follow Up code - show the Employee Flag record
'                Call Show_Employee_Flag(x)
'            End If
'        Else
'            If Len(lblTitle(x).Caption) > 0 Then
'                'User has right to access this Follow Up code - show the Employee Flag record
'                Call Show_Employee_Flag(x)
'            End If
'        End If
'    Next
    
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
UpdateRight = gSec_Upd_EMP_FLAGS
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
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call lockctl(Me, TF)
End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEMPLOYEEFLAGS" Then glbOnTop = ""

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdModify_Click()

On Error GoTo Mod_Err

'oEmail = txtEEMail(0)

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HREMP_FLAG", "Modify")
Call RollBack

End Sub

Public Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

rsDATA.CancelUpdate

Call Display_Value

fglbNew = False

Call SET_UP_MODE

'Based on Follow Up Code security
Call HideShow_EmployeeFlagsRec_FollowUp_Security

'For x = 0 To 19
'    'Call procedure to check if user has access to the follow up code
'    If Len(clpCode(x).Text) > 0 And clpCode(x).Caption <> "Unassigned" And clpCode(x).Caption <> "" Then
'        If Not Accessible_FollowUp_Code(clpCode(x).Text) Then
'            'User does not have right to access this Follow Up code - hide the Employee Flag record
'            Call Hide_Employee_Flag(x)
'        Else
'            'User has right to access this Follow Up code - show the Employee Flag record
'            Call Show_Employee_Flag(x)
'        End If
'    Else
'        If Len(lblTitle(x).Caption) > 0 Then
'            'User has right to access this Follow Up code - show the Employee Flag record
'            Call Show_Employee_Flag(x)
'        End If
'    End If
'Next


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack   '23June99 js

End Sub

Public Sub cmdOK_Click()

On Error GoTo Add_Err
Dim X As Long, vList As String

If chkEmpFlags = False Then Exit Sub

'rsDATA.Requery
Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)
If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    
    If gsAttachment_DB Then
        For X = 0 To lblTitle.count - 1
            'lblTitle(X).Caption = lStr(lblTitle(X).Caption)
            If Not IsNull(lblTitle(X).Caption) And lblTitle(X).Caption <> "" Then
                gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_EMP_FLAGS set EF_FLAGDTE=" & Date_SQL(rsDATA("EF_FLAGDTE" & X + 1)) & " WHERE EF_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " AND EF_FLAG = " & X '& " AND EF_FLAGDTE=" & Date_SQL(oldFlagDt(X))
            End If
        Next
    End If
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    
    If gsAttachment_DB Then
        For X = 0 To lblTitle.count - 1
            'lblTitle(X).Caption = lStr(lblTitle(X).Caption)
            If Not IsNull(lblTitle(X).Caption) And lblTitle(X).Caption <> "" Then
                gdbAdoIhr001_DOC.Execute "Update HRDOC_EMP_FLAGS set EF_FLAGDTE=" & Date_SQL(rsDATA("EF_FLAGDTE" & X + 1)) & " WHERE EF_TYPE='" & UCase(glbDocName) & "' AND EF_EMPNBR = " & glbLEE_ID & " AND EF_FLAG = " & X '& " AND EF_FLAGDTE=" & Date_SQL(oldFlagDt(X))
            End If
        Next
    End If
End If
Data1.Refresh

For X = 0 To cboEmpFlag.count - 1
    vList = FlgList(X)
    If Len(dlpFU(X).Text) > 0 Then
        Call updFollow("U", X)
    End If
Next

fglbNew = False

Call SET_UP_MODE

Call EERetrieve

Call NextForm

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP_FLAGS", "Update")
Call RollBack  '23June99 - js

End Sub

Private Sub scrHorz_Change()
    panDetails.Left = 0 - scrHorz.Value
End Sub

Private Sub txtEmpFlag_Change(Index As Integer)
    cboEmpFlag(Index).Text = Trim(txtEmpFlag(Index).Text)
End Sub

Private Function FlgList(idx As Long) As String
Dim retVal As String, FlgFile
retVal = ""
FlgFile = glbIHRREPORTS & "FlagList" & CStr(idx) & ".MTF"

On Error GoTo ErrorHandler

If File(FlgFile) Then
    Open FlgFile For Input As #1
    Input #1, retVal
    Close #1
End If

ResumeHere:
If InStr(retVal, Trim(cboEmpFlag(idx))) = 0 And Trim(cboEmpFlag(idx)) <> "" Then
    retVal = retVal & "&" & Trim(cboEmpFlag(idx))
    cboEmpFlag(idx).AddItem Trim(cboEmpFlag(idx))
End If
Open FlgFile For Output As #1
Print #1, retVal
Close #1
FlgList = retVal
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt FlagList" & CStr(idx) & ".MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill FlgFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub scrControl_Change()
    panDetails.Top = 0 - scrControl.Value
End Sub

Private Function chkEmpFlags() As Boolean
    On Error GoTo Eh
    Dim C As Long
    Dim rs As New ADODB.Recordset
    Dim SQLQ As String
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    
    chkEmpFlags = False
    
    For C = 0 To 19
        If Len(Trim(cboEmpFlag(C).Text)) > 25 Then
            MsgBox "'" & lblTitle(C).Caption & "' value cannot be greater than 25 characters", vbInformation + vbOKOnly, "Maximum length of Value"
            cboEmpFlag(C).SetFocus
            Exit Function
        End If
            
        If Len(dlpFU(C).Text) > 0 And IsDate(dlpFU(C).Text) = False Then
            MsgBox "'" & lblTitle(C).Caption & lStr("' Follow-Up date must be a valid date"), vbInformation + vbOKOnly, "Invalid Reason code"
            dlpFU(C).SetFocus
            Exit Function
        End If
        
        If Len(dlpFU(C).Text) > 0 And Len(clpCode(C).Text) <= 0 Then
            MsgBox "'" & lblTitle(C).Caption & lStr("' Follow-Up reason cannot be blank"), vbInformation + vbOKOnly, "Invalid Reason code"
            clpCode(C).SetFocus
            Exit Function
        End If
        
        If clpCode(C).Caption = "Unassigned" Then
            MsgBox "'" & lblTitle(C).Caption & lStr("' Follow-Up reason must be valid"), vbInformation + vbOKOnly, "Invalid Reason code"
            clpCode(C).SetFocus
            Exit Function
        End If
        
        If Len(dlpFU(C).Text) > 0 And clpCode(C).Caption = "Unassigned" Then
            MsgBox "'" & lblTitle(C).Caption & lStr("' Follow-Up reason must be valid"), vbInformation + vbOKOnly, "Invalid Reason code"
            clpCode(C).SetFocus
            Exit Function
        End If
        
        
        If Len(clpCode(C).Text) > 0 And clpCode(C).Visible = True And clpCode(C).Enabled = True Then
            'Check to see if valid Follow Up code as per the security
            If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
            Else
                '????Ticket #24808 -  Retrieve template's security profile
                SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            End If
            'SQLQ = "SELECT ACCESSABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & glbUserID & "'"
            SQLQ = SQLQ & " AND CODENAME='" & clpCode(C).Text & "'"
            rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                If rs("MAINTAINABLE") = 0 Then
                'If rs("ACCESSABLE") = 0 Then
                    MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(C).Text & "' Reason code.", vbOKOnly + vbInformation, "Authorization failed"
                    rs.Close
                    Set rs = Nothing
                    'clpCode(C).SetFocus
                    Exit Function
                End If
            Else
                MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(C).Text & "' Reason code.", vbOKOnly + vbInformation, "Authorization failed"
                rs.Close
                Set rs = Nothing
                clpCode(C).SetFocus
                Exit Function
            End If
            rs.Close
            Set rs = Nothing
        End If
        
        If Len(dlpFU(C).Text) = 0 And Len(clpCode(C).Text) > 0 Then
            MsgBox "'" & lblTitle(C).Caption & lStr("' Follow-Up Date cannot be blank"), vbInformation + vbOKOnly, "Invalid " & lStr("Follow-Up Date")
            dlpFU(C).SetFocus
            Exit Function
        End If
        
    Next C
    chkEmpFlags = True
    
exH:
    Exit Function
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP_FLAGS", "Update")
    Call RollBack  '23June99 - js
    Resume exH
End Function


Private Function updFollow(xType As String, idx As Long) 'created by Bryan 29/Mar/06
Dim newline As String
Dim SQLQ As String
Dim Msg As String, Edat As String
Dim iRec As Integer
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Boolean
'Don't need a message for follow up - Jerry asked for v7.6

updFollow = False


On Error GoTo CrFollow_Err

newline = Chr$(13) & Chr$(10)

If IsDate(dlpFU(idx).Text) Then 'Jaddy 11/15
    ' DATE Renewal IS NOW MANDATORY
    If glbtermopen Then
        SQLQ = "SELECT * FROM Term_FOLLOW_UP WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND EF_FREAS = '" & clpCode(idx).Text & "'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpFU(idx).Text)
        'SQLQ = SQLQ & " AND EF_COMMENTS = '" & lblTitle(idx).Caption & " - " & cboEmpFlag(idx).Text & "'"
        dynHRAT.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & Val(glbLEE_ID)
        SQLQ = SQLQ & " AND EF_FREAS = '" & clpCode(idx).Text & "'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpFU(idx).Text)
        'SQLQ = SQLQ & " AND EF_COMMENTS = '" & lblTitle(idx).Caption & " - " & cboEmpFlag(idx).Text & "'"
        dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" And Edit1 = False Then
    
    If glbtermopen Then
        rsTB.Open "Term_FOLLOW_UP", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    Else
        rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    End If
    If IsDate(dlpFU(idx).Text) And Edit1 = False Then 'Jaddy 11/15
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        If glbtermopen Then
            rsTB("EF_EMPNBR") = glbTERM_ID
            rsTB("TERM_SEQ") = glbTERM_Seq
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY") = GetTermEmpData(glbTERM_ID, glbTERM_Seq, "ED_ADMINBY", Null)
            End If
        Else
            rsTB("EF_EMPNBR") = glbLEE_ID
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
        End If
        rsTB("EF_FDATE") = CVDate(dlpFU(idx).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
        End If
        rsTB("EF_COMMENTS") = lblTitle(idx).Caption & " - " & Trim(cboEmpFlag(idx).Text)
        rsTB("EF_FREAS") = clpCode(idx).Text
        rsTB("EF_LDATE") = Updstats(0).Text
        rsTB("EF_LTIME") = Updstats(1).Text
        rsTB("EF_LUSER") = Updstats(2).Text
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = lStr("A Follow-Up Record was created!")
        'MsgBox Msg
        Exit Function
    End If
    
    dynHRAT.Close
    updFollow = True
    Edit1 = True
    Exit Function

Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            'Ticket #17916
            'Not sure what to do here but this part was called whenever Edit = True which did not make
            'sense as it was deleting an existing record - so I have commented it out.
            'If a user deletes the followup date/reason then they will have to manually delete the corresponding
            'Follow Up record
            'dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        'Ticket #17916 - commented it out as no delete had taken place
        'Msg = lStr("A record has been deleted from the Follow-Up table")
        
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function

Private Sub Hide_Employee_Flag(xInd)
    'lblTitle(xInd).Visible = False
    cboEmpFlag(xInd).Visible = False
    dlpDate(xInd).Visible = False
    dlpFU(xInd).Visible = False
    clpCode(xInd).Visible = False
End Sub

Private Sub Show_Employee_Flag(xInd)
    lblTitle(xInd).Visible = True
    cboEmpFlag(xInd).Visible = True
    dlpDate(xInd).Visible = True
    dlpFU(xInd).Visible = True
    clpCode(xInd).Visible = True
    
    'Check if the user has Maintain right on the Follow Up Code to be able to make the changes to
    'the existing record
    If Len(clpCode(xInd).Text) > 0 And clpCode(xInd).Caption <> "Unassigned" And clpCode(xInd).Caption <> "" Then
        If Not Maintainable_FollowUp_Code(clpCode(xInd).Text) Then
            dlpFU(xInd).Enabled = False
            clpCode(xInd).Enabled = False
        Else
            dlpFU(xInd).Enabled = True
            clpCode(xInd).Enabled = True
        End If
    Else
        dlpFU(xInd).Enabled = True
        clpCode(xInd).Enabled = True
    End If
    
End Sub

Private Function Accessible_FollowUp_Code(xFUCode) As Boolean
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)

    'SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & glbUserID & "'"
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    SQLQ = SQLQ & " AND CODENAME='" & xFUCode & "'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        'If rs("MAINTAINABLE") = 0 Then
        If rs("ACCESSABLE") = 0 Then
            Accessible_FollowUp_Code = False
        Else
            Accessible_FollowUp_Code = True
        End If
    Else
        Accessible_FollowUp_Code = False
    End If
    rs.Close
    Set rs = Nothing

End Function

Private Function Maintainable_FollowUp_Code(xFUCode) As Boolean
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)

    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    'SQLQ = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & glbUserID & "'"
    SQLQ = SQLQ & " AND CODENAME='" & xFUCode & "'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rs.EOF Then
        If rs("MAINTAINABLE") = 0 Then
        'If rs("ACCESSABLE") = 0 Then
            Maintainable_FollowUp_Code = False
        Else
            Maintainable_FollowUp_Code = True
        End If
    Else
        Maintainable_FollowUp_Code = False
    End If
    rs.Close
    Set rs = Nothing

End Function

Private Function FollowUp_Sec() As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As Boolean
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    
    strSQL = "SELECT MAINTAINABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    'strSQL = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        retVal = Abs(rs("MAINTAINABLE"))
        'retVal = Abs(rs("ACCESSABLE"))
    Else
        retVal = False
    End If
    
    FollowUp_Sec = retVal
End Function

Private Sub HideShow_EmployeeFlagsRec_FollowUp_Security()
    Dim X As Integer
    
    'Based on Follow Up Code security

    For X = 0 To 19
        'Call procedure to check if user has access to the follow up code
        If Len(clpCode(X).Text) > 0 And clpCode(X).Caption <> "Unassigned" And clpCode(X).Caption <> "" Then
            If Not Accessible_FollowUp_Code(clpCode(X).Text) Then
                'User does not have right to access this Follow Up code - hide the Employee Flag record
                Call Hide_Employee_Flag(X)
            Else
                'User has right to access this Follow Up code - show the Employee Flag record
                Call Show_Employee_Flag(X)
            End If
        Else
            If Len(lblTitle(X).Caption) > 0 Then
                'User has right to access this Follow Up code - show the Employee Flag record
                Call Show_Employee_Flag(X)
            End If
        End If
    Next

End Sub
