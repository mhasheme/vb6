VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmQuickESS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Quick Setup for ESS"
   ClientHeight    =   8190
   ClientLeft      =   105
   ClientTop       =   750
   ClientWidth     =   10350
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10065.79
   ScaleMode       =   0  'User
   ScaleWidth      =   18427.46
   Tag             =   "Close and exit this screen."
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   10350
      _Version        =   65536
      _ExtentX        =   18256
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Enabled         =   0   'False
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "USERID"
         DataField       =   "USERID"
         DataSource      =   "Data1"
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
         Left            =   750
         TabIndex        =   26
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
         DataField       =   "USERName"
         DataSource      =   "Data1"
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
         Left            =   3045
         TabIndex        =   25
         Top             =   120
         Width           =   1290
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   2850
      Left            =   0
      TabIndex        =   27
      Top             =   495
      Width           =   10350
      _Version        =   65536
      _ExtentX        =   18256
      _ExtentY        =   5027
      _StockProps     =   15
      Caption         =   "SSPanel1"
      ForeColor       =   12632256
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
      Begin VB.TextBox txtEEName 
         Appearance      =   0  'Flat
         DataField       =   "USERNAME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   3
         Text            =   "txtEEID"
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtUSERID 
         Appearance      =   0  'Flat
         DataField       =   "USERID"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1320
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "00-User ID"
         Text            =   "txtUSERID"
         Top             =   2520
         Width           =   1305
      End
      Begin VB.TextBox txtEEID 
         Appearance      =   0  'Flat
         DataField       =   "EMPNBR"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   4440
         TabIndex        =   2
         Tag             =   "10-Employee #"
         Text            =   "txtEEID"
         Top             =   2520
         Width           =   1095
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "frmQuickESS.frx":0000
         Height          =   2370
         Left            =   120
         Negotiate       =   -1  'True
         OleObjectBlob   =   "frmQuickESS.frx":0014
         TabIndex        =   0
         Tag             =   "Listing of Security Records"
         Top             =   0
         Width           =   9015
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   195
         Left            =   5880
         TabIndex        =   31
         Top             =   2610
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
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
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2610
         Width           =   660
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         DataField       =   "EMPNBR"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3960
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Height          =   195
         Left            =   3000
         TabIndex        =   28
         Top             =   2610
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panDetails 
      Height          =   4695
      Index           =   0
      Left            =   0
      TabIndex        =   32
      Top             =   3360
      Width           =   11805
      _Version        =   65536
      _ExtentX        =   20823
      _ExtentY        =   8281
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
      Begin VB.TextBox txtSecPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         DataField       =   "PassWord"
         DataSource      =   "Data1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   44
         Tag             =   "00-Password"
         Top             =   840
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox txtConfPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   42
         Tag             =   "00-Confirm Password"
         Top             =   600
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "Grant All for Quick Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5640
         TabIndex        =   23
         Tag             =   "Grant All for Quick Setup"
         Top             =   2965
         Width           =   2385
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "LDATE"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   8160
         MaxLength       =   25
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "LTIME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   7560
         MaxLength       =   25
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "LUSER"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   7560
         MaxLength       =   25
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtPWord 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Tag             =   "00-Password"
         Top             =   240
         Width           =   1305
      End
      Begin VB.ListBox lstDept 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4440
         TabIndex        =   5
         Tag             =   "00-Department(s)"
         Top             =   120
         Width           =   1095
      End
      Begin Threed.SSCheck chkRCompT 
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5640
         TabIndex        =   20
         Top             =   1800
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Compensatory Time"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkEESecurity 
         DataField       =   "EmpNBR_Based"
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5640
         TabIndex        =   21
         Top             =   2040
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Employee Number Based Security"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkEESIN 
         DataSource      =   "Data1"
         Height          =   315
         Left            =   5640
         TabIndex        =   22
         Top             =   2280
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   556
         _StockProps     =   78
         Caption         =   "Show SIN/SSN "
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AAF4
         DataSource      =   "Data1"
         Height          =   225
         Index           =   5
         Left            =   1560
         TabIndex        =   17
         Top             =   2960
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Hourly Entitlements"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Bindings        =   "frmQuickESS.frx":AAFF
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   2960
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Bindings        =   "frmQuickESS.frx":AB0A
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   2740
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Bindings        =   "frmQuickESS.frx":AB15
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   10
         Top             =   2300
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   2080
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB20
         DataSource      =   "Data1"
         Height          =   225
         Index           =   4
         Left            =   1560
         TabIndex        =   15
         Top             =   2740
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Attendance"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB2B
         DataSource      =   "Data1"
         Height          =   225
         Index           =   3
         Left            =   1560
         TabIndex        =   13
         Top             =   2520
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Sick/Vacation Entitlements"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB36
         DataSource      =   "Data1"
         Height          =   225
         Index           =   2
         Left            =   1560
         TabIndex        =   11
         Top             =   2300
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Benefits"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB41
         DataSource      =   "Data1"
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   2080
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Dependents"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB4C
         DataSource      =   "Data1"
         Height          =   225
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   1860
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Employee Demographics / Dates"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1860
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Bindings        =   "frmQuickESS.frx":AB57
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   18
         Top             =   3180
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB62
         DataSource      =   "Data1"
         Height          =   225
         Index           =   6
         Left            =   1560
         TabIndex        =   19
         Top             =   3180
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Counselling"
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
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Bindings        =   "frmQuickESS.frx":AB6D
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   45
         Top             =   3400
         Width           =   435
         _Version        =   65536
         _ExtentX        =   767
         _ExtentY        =   344
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmQuickESS.frx":AB78
         DataSource      =   "Data1"
         Height          =   225
         Index           =   7
         Left            =   1560
         TabIndex        =   46
         Top             =   3400
         Width           =   3105
         _Version        =   65536
         _ExtentX        =   5477
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Archive Vacation/Timeoff"
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
         Font3D          =   3
      End
      Begin VB.Label lblConfPass 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Conf.Pass"
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
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   630
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
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
         Height          =   195
         Index           =   0
         Left            =   1560
         TabIndex        =   41
         Top             =   1320
         Width           =   600
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
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
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   40
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblCNum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label20"
         DataField       =   "COMPNO"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   6480
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label1"
         DataField       =   "ID"
         DataSource      =   "Data1"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   8640
         TabIndex        =   38
         Top             =   3240
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblPWord 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblDeps 
         Caption         =   "Department(s)"
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
         Left            =   3000
         TabIndex        =   36
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   405
      Left            =   2520
      Top             =   7800
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   714
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   1200
      Top             =   7920
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
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_File_ESecurity 
         Caption         =   "Exit Quick Security for ESS"
      End
      Begin VB.Menu mnu_F_Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "Exit INFO:HR"
      End
   End
End
Attribute VB_Name = "frmQuickESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew%, fglbView%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim OUserID, ODept 'keep current state of database in case when user wants to edit
Dim x%
Dim ChangeCBox
Dim tPass As String
Dim ChkPass
Dim Qu


Private Sub chkMSecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'If chkSecurity(Index).Value = False Then
'    chkMSecurity(Index).Value = False
'End If
If chkMSecurity(Index).Value = True Then
    chkSecurity(Index).Value = True
End If

End Sub

Private Sub chkMSecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
'If chkSecurity(Index).Value Then
'    chkMSecurity(Index).Value = False
'End If
If chkMSecurity(Index).Value = True Then
    chkSecurity(Index).Value = True
End If

End Sub

Private Sub chkSecurity_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If chkSecurity(Index).Value = False Then
    chkMSecurity(Index).Value = False
End If

End Sub

Private Sub chkSecurity_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If chkSecurity(Index).Value = False Then
    chkMSecurity(Index).Value = False
End If

End Sub

Private Sub cmdGrantAll_Click()
Dim x%

For x% = 0 To 7
    chkMSecurity(x%).Value = True
    chkSecurity(x%).Value = True
Next x%

chkRCompT.Value = True
chkEESecurity.Value = True
chkEESIN.Value = True

End Sub

Private Sub cmdGrantAll_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbFNo = 1
End Sub

Private Sub Form_GotFocus()
    Dim SQLE
    Dim rsEMP As New ADODB.Recordset
    Data1.Refresh

    SQLE = "SELECT PD_DEPT FROM HRPASDEP WHERE PD_USERID = '" & Replace(txtUSERID, "'", "''") & "'"
    rsEMP.Open SQLE, gdbAdoIhr001, adOpenStatic

    lstDept.Clear
    Do While Not rsEMP.EOF
        lstDept.AddItem (rsEMP("PD_DEPT"))
        rsEMP.MoveNext
    Loop
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMQUICKESS"
Dim Answer, DefVal, Msg, Title, x%  '  variables.
Dim RFound As Integer ' records found
Dim SQLQ, SQLE
On Error GoTo SecureLoad_Err
Dim rsEMP As New ADODB.Recordset

Screen.MousePointer = HOURGLASS
fglbNew% = False
glbFNo = 1
Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT *, "
If glbLinamar Then
    SQLQ = SQLQ & " CASE WHEN EMPNBR IS NOT NULL AND LEN(EMPNBR)>2 "
    SQLQ = SQLQ & " THEN RIGHT(EMPNBR,3)+'-'+"
    SQLQ = SQLQ & " LEFT(EMPNBR,LEN(EMPNBR)-3) "
    SQLQ = SQLQ & " ELSE STR(EMPNBR) END "
    SQLQ = SQLQ & " AS SEMPNBR "
Else
    SQLQ = "SELECT * "
    vbxTrueGrid.Columns(1).DataField = "EMPNBR"
End If
SQLQ = SQLQ & " from HR_SECURE_BASIC ORDER BY USERID"
Data1.RecordSource = SQLQ
Data1.Refresh

SQLE = "SELECT PD_DEPT FROM HRPASDEP WHERE PD_USERID = '" & Replace(txtUSERID, "'", "''") & "'"
rsEMP.Open SQLE, gdbAdoIhr001, adOpenStatic

lstDept.Clear
Do While Not rsEMP.EOF
    lstDept.AddItem (rsEMP("PD_DEPT"))
    rsEMP.MoveNext
Loop

Call mod_UpdateMode(False)

If Not gSec_Upd_Security Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
'    cmdNew.Enabled = False '
End If                                                  '
fglbView% = 0
panDetails(fglbView%).Visible = True
Screen.MousePointer = DEFAULT
Exit Sub

SecureLoad_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Load", "HR_SECURE", "Select")
Call RollBack

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim VR
Call ChkCBoxChange
If ChangeCBox = True Then
        VR = MsgBox("Do you want to save changes?", MB_YESNO)
        If VR = IDYES Then
            Me.cmdOK_Click 'Then Pause (0.5) Else isUpdated = False
        ElseIf VR = IDNO Then
            Call Me.cmdCancel_Click
        End If
End If

If Not IsNull(Data1.Recordset("PassWord")) Then
    If gsMultiLang = "YES" Then
        If txtPWord.Text <> DecryptPasswordMultiLang(Data1.Recordset("PassWord")) Then
             glbConfPass = txtPWord
            Load frmConfPass 'Call ChkPassUnl
            frmConfPass.Show vbModal
        End If
    Else
        If txtPWord.Text <> DecryptPassword(Data1.Recordset("PassWord")) Then
             glbConfPass = txtPWord
            Load frmConfPass 'Call ChkPassUnl
            frmConfPass.Show vbModal
        End If
    End If
End If

Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."

End Sub

Private Sub lblEEID_Change()
txtEEID = ShowEmpnbr(lblEEID)
If Data1.Recordset.EOF Then Exit Sub
End Sub


Private Sub lstDept_DblClick()

If Len(txtUSERID) <= 0 Then
    MsgBox "User ID is required"
Else

glbSecUSERID = txtUSERID

If Len(lblEEName.Caption) > 0 Then
    glbSecEEName$ = lblEEName.Caption
Else
    glbSecEEName$ = " "
End If

frmSDept.Show

End If

End Sub

Private Sub mnu_File_ESecurity_Click()
Unload Me
End Sub

Private Sub txtEEID_Change()
If Len(txtEEID) > 0 Then txtEEName.Enabled = False Else txtEEName.Enabled = True 'And cmdOK.Enabled
End Sub

Private Sub txtEEID_DblClick()
Dim LastlastID As Long, LastlastNme As String, LastFirstNme As String

LastlastNme = glbLEE_SName
LastFirstNme = glbLEE_FName
LastlastID = glbLEE_ID

frmEEFIND.Show 1

If glbLEE_ID <> 0 Then
    If glbEEOK Then 'don't do recall unless different from this
        txtEEID = ShowEmpnbr(glbLEE_ID)
        txtEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
End If

glbLEE_SName = LastlastNme
glbLEE_FName = LastFirstNme
glbLEE_ID = LastlastID

End Sub

Private Sub txtEEID_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEEID_KeyPress(KeyAscii As Integer)
    If glbLinamar And KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then KeyAscii = 0
End Sub

Private Sub txtEEID_LostFocus()
Dim rsEMP As New ADODB.Recordset
Dim SQLQ
If Len(txtEEID) > 0 Then
    SQLQ = "SELECT ED_SURNAME,ED_FNAME FROM HREMP "
    SQLQ = SQLQ & "Where ED_EMPNBR = " & getEmpnbr(txtEEID)

    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not (rsEMP.BOF Or rsEMP.EOF) Then
        txtEEName = rsEMP("ED_SURNAME") & ", " & rsEMP("ED_FNAME")
    End If
    rsEMP.Close
End If
End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
Dim x%

On Error GoTo Mod_Err
OUserID = txtUSERID
Call SET_UP_MODE
'Call mod_UpdateMode(True)

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack

End Sub

Sub cmdOK_Click()
Dim xID

Dim OPwd As String

If Not chkSecureOk() Then Exit Sub

If IsNull(Data1.Recordset("PassWord")) Then
    OPwd = ""
Else
    OPwd = Data1.Recordset("PassWord")
End If

If gsMultiLang = "YES" Then
    If txtPWord.Text <> DecryptPasswordMultiLang(OPwd) Then
        glbConfPass = txtPWord
        Load frmConfPass
        frmConfPass.Show vbModal
        If glbConfPass = "" Then
            Exit Sub
        End If
    End If
Else
    If txtPWord.Text <> DecryptPassword(OPwd) Then
        glbConfPass = txtPWord
        Load frmConfPass
        frmConfPass.Show vbModal
        If glbConfPass = "" Then
            Exit Sub
        End If
    End If
End If

On Error GoTo Add_Err

'If Not chkSecureOk() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

panEEDESC.Enabled = False

xID = txtUSERID
If Len(txtEEID) = 0 Then
    Data1.Recordset("EMPNBR") = Null
Else
    Data1.Recordset("EMPNBR") = getEmpnbr(txtEEID)
End If
Data1.Recordset("USERID") = xID


Data1.Recordset.UpdateBatch

If fglbNew Then
    Call AddSecAccess
Else
    Call UpdSecAccess
    If OUserID <> txtUSERID Then Call UpdateRelated
End If
If Not glbSQL And Not glbOracle Then Pause (0.5)
Data1.Refresh
Data1.Recordset.Find "USERID='" & Replace(xID, "'", "''") & "'"
glbChkPass = False
fglbNew% = False

'Call mod_UpdateMode(False)

Call SET_UP_MODE

Exit Sub
Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OCC_HEALTH_SAFETY", "Update")
Call RollBack   '10June99 js
End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh



fglbNew% = False
Call SET_UP_MODE


panEEDESC.Enabled = False

panDetails(fglbView%).Visible = True

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack

End Sub


Sub cmdNew_Click()
Dim SQLQ As String

Dim VR
fglbNew% = True
Call ChkCBoxChange
If ChangeCBox = True Then
        VR = MsgBox("Do you want to save changes?", MB_YESNO)
        If VR = IDYES Then
            Me.cmdOK_Click 'Then Pause (0.5) Else isUpdated = False
        ElseIf VR = IDNO Then
            Call Me.cmdCancel_Click
        End If
End If

fglbNew% = True

panEEDESC.Enabled = True
ODept = ""

Call SET_UP_MODE
On Error GoTo AddN_Err


Data1.Recordset.AddNew

lstDept.AddItem "DoubleClick"
lblCNum.Caption = "001"
txtUSERID.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OCC_HEALTH_SAFETY", "Add")
Call RollBack

End Sub
Sub cmdDelete_Click()
Dim a As Integer, Msg$, INo&, x%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

If lblEEID = "999999999" Then
    Msg$ = "You can not delete the master security record"
    Msg$ = Msg$ & Chr(10) & "You can however, change its password."
    MsgBox Msg$
    Exit Sub
End If

On Error GoTo Del_Err


Msg$ = "Are You Sure You Want To Delete "
Msg$ = Msg$ & Chr(10) & "This Record?  "

a% = MsgBox(Msg$, 36, "Confirm Delete")
'If user press Yes button information that were inserted will be deleted
If a% <> 6 Then Exit Sub
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_BASIC WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(lblUSERID, "'", "''") & "'"
gdbAdoIhr001.Execute "DELETE FROM HRPASDEP WHERE PD_USERID ='" & Replace(lblUSERID, "'", "''") & "'"

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
fglbNew = False
Call SET_UP_MODE

'Call mod_UpdateMode(False)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
Call RollBack

End Sub


Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%
'cmdPrint.Enabled = False
Me.vbxCrystal.WindowTitle = "Security Master Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 4
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
        Me.vbxCrystal.DataFiles(6) = glbIHRDB
        Me.vbxCrystal.DataFiles(7) = glbIHRDB
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGSECURE.rpt"
    Call SECWRK
'    Me.vbxCrystal.SelectionFormula = "{HR_SECURE_BASIC.USERID}='" & txtUSERID & "' "
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, x%
'cmdPrint.Enabled = False
Me.vbxCrystal.WindowTitle = "Security Master Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 4
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
        Me.vbxCrystal.DataFiles(6) = glbIHRDB
        Me.vbxCrystal.DataFiles(7) = glbIHRDB
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGSECURE.rpt"
    Call SECWRK
'    Me.vbxCrystal.SelectionFormula = "{HR_SECURE_BASIC.USERID}='" & txtUSERID & "' "
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub


Private Sub mnu_File_ESecure_Click()
    Unload Me
End Sub

Private Sub mnu_File_Exit_Click()
End
End Sub


Private Sub mod_UpdateMode(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

fUPMode = TF    ' update mode
'Me.vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
panDetails(0).Enabled = TF      '10June99 js
'mnu_File.Enabled = FT
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
txtUSERID.Enabled = TF
txtEEID.Enabled = TF
'txtEEName.Enabled = TF
If Data1.Recordset.BOF And Data1.Recordset.EOF Then '
'   cmdModify.Enabled = False
'   cmdDelete.Enabled = False
End If                                              '

End Sub

Private Function displaypanel()

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

Private Function chkSecureOk()
Dim rsEMP As New ADODB.Recordset
Dim SQLQ As String, Msg$, ID As Long
Dim snapSec As New ADODB.Recordset

Screen.MousePointer = HOURGLASS
chkSecureOk = False

On Error GoTo chkSecureOk_Err
If Len(txtUSERID) <= 0 Then
    MsgBox "User ID is required"
    txtUSERID.SetFocus
    GoTo chkExit
End If
If Len(txtEEName) <= 0 Then
    MsgBox "User Name is required"
    txtEEName.SetFocus
    GoTo chkExit
End If

ID& = Val(lblID.Caption)
SQLQ = "SELECT * FROM HR_SECURE_BASIC "
SQLQ = SQLQ & "Where (USERID = '" & Replace(txtUSERID, "'", "''") & "'"
SQLQ = SQLQ & " AND ID <> " & ID & ") "

snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapSec.BOF And snapSec.EOF Then
   Rem everything is ok
Else
    MsgBox "This user already has a security record"
    GoTo chkExit
End If
snapSec.Close
If Len(txtEEID) > 0 Then
    SQLQ = "SELECT ED_EMPNBR FROM HREMP "
    SQLQ = SQLQ & "Where ED_EMPNBR = " & getEmpnbr(txtEEID)

    snapSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapSec.BOF And snapSec.EOF Then 'not a valid ee
        MsgBox "This employee does not exist"
        GoTo chkExit
    End If
    snapSec.Close
End If

If Len(txtPWord) < 1 Or Len(txtPWord) > 15 Then
    MsgBox "Invalid Password (must be between 1 and 15 characters)'"
    txtPWord.SetFocus
    GoTo chkExit
End If

SQLQ = "SELECT PD_DEPT FROM HRPASDEP WHERE PD_USERID = '" & Replace(txtUSERID, "'", "''") & "'"
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
       
If rsEMP.EOF Then
    MsgBox "Department is required"
    GoTo chkExit
End If


chkSecureOk = True

chkExit:
Screen.MousePointer = DEFAULT
Exit Function

chkSecureOk_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OCC_HEALTH_SAFETY", "edit/Add")
Call RollBack   '10June99 js

End Function
Private Function panVisible()

End Function
Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim x%, SQLQ

SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE "
SQLQ = SQLQ & " FROM HR_SECURE_ACCESS "
SQLQ = SQLQ & " WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND Maintainable=0"
rsSR.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
Call ResetAll
Do Until rsSR.EOF
    
    If UCase(rsSR("FUNCTION")) = UCase("Basic_Update") Then chkMSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Basic_Inquiry") Then chkSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Dependents_Update") Then chkMSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Dependents_Inquiry") Then chkSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Benefits_Update") Then chkMSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Benefits_Inquiry") Then chkSecurity(2) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Entitlements_Update") Then chkMSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Entitlements_Inquiry") Then chkSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_Update") Then chkMSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Attendance_Inquiry") Then chkSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Hrly_Entitlements_Update") Then chkMSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Hrly_Entitlements_Inquiry") Then chkSecurity(5) = rsSR("ACCESSABLE")
    
    If UCase(rsSR("FUNCTION")) = UCase("Counselling_Update") Then chkMSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Counselling_Inquiry") Then chkSecurity(6) = rsSR("ACCESSABLE")
    ' Sam 07/27/2006 Ticket # 11043
    If UCase(rsSR("FUNCTION")) = UCase("Archive_VacTimeoff_Update") Then chkMSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Archive_VacTimeoff_Inquiry") Then chkSecurity(7) = rsSR("ACCESSABLE")
    'Ends
    If UCase(rsSR("FUNCTION")) = UCase("Show_SIN_SSN") Then chkEESIN = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("Report_Compensatory_Time") Then chkRCompT = rsSR("ACCESSABLE")
    
    rsSR.MoveNext
Loop
Call SET_UP_MODE
Me.cmdModify_Click
End Sub
Private Sub AddSecAccess()
Dim SQLQ, sqlI
sqlI = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID," & Field_SQL("FUNCTION") & ",ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(txtUSERID), "'", "''") & "',"

SQLQ = sqlI & "'Basic_Update'," & IIf(chkMSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Basic_Inquiry'," & IIf(chkSecurity(0), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Dependents_Update'," & IIf(chkMSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Dependents_Inquiry'," & IIf(chkSecurity(1), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Benefits_Update'," & IIf(chkMSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Benefits_Inquiry'," & IIf(chkSecurity(2), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Entitlements_Update'," & IIf(chkMSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Entitlements_Inquiry'," & IIf(chkSecurity(3), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Attendance_Update'," & IIf(chkMSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Attendance_Inquiry'," & IIf(chkSecurity(4), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 SQLQ = sqlI & "'Hrly_Entitlements_Update'," & IIf(chkMSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Hrly_Entitlements_Inquiry'," & IIf(chkSecurity(5), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Counselling_Update'," & IIf(chkMSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Counselling_Inquiry'," & IIf(chkSecurity(6), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 
 'sam added 07/27/2006 Ticket # 11043
SQLQ = sqlI & "'Archive_VacTimeoff_Update'," & IIf(chkMSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Archive_VacTimeoff_Inquiry'," & IIf(chkSecurity(7), 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ

'ends
 
SQLQ = sqlI & "'Show_SIN_SSN'," & IIf(chkEESIN, 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'Report_Compensatory_Time'," & IIf(chkRCompT, 1, 0) & ")"
 gdbAdoIhr001.Execute SQLQ
 

End Sub
Private Sub UpdSecAccess()
Dim SQLQ
SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND CODENAME is NULL"
SQLQ = SQLQ & " and " & Field_SQL("FUNCTION") & " IN ("
SQLQ = SQLQ & "'Basic_Update',"
SQLQ = SQLQ & "'Basic_Inquiry',"
SQLQ = SQLQ & "'Dependents_Update',"
SQLQ = SQLQ & "'Dependents_Inquiry',"
SQLQ = SQLQ & "'Benefits_Update',"
SQLQ = SQLQ & "'Benefits_Inquiry',"
SQLQ = SQLQ & "'Entitlements_Update',"
SQLQ = SQLQ & "'Entitlements_Inquiry',"
SQLQ = SQLQ & "'Attendance_Update',"
SQLQ = SQLQ & "'Attendance_Inquiry',"
SQLQ = SQLQ & "'Hrly_Entitlements_Update',"
SQLQ = SQLQ & "'Hrly_Entitlements_Inquiry',"
SQLQ = SQLQ & "'Counselling_Update',"
SQLQ = SQLQ & "'Counselling_Inquiry',"
'sam added 07/27/2006 Ticket # 11043
SQLQ = SQLQ & "'Archive_VacTimeoff_Update',"
SQLQ = SQLQ & "'Archive_VacTimeoff_Inquiry',"
'ends
SQLQ = SQLQ & "'Show_SIN_SSN',"
SQLQ = SQLQ & "'Report_Compensatory_Time'"
SQLQ = SQLQ & ")"

gdbAdoIhr001.Execute SQLQ

Call AddSecAccess
End Sub
Private Sub UpdateRelated()
Dim SQLQ
SQLQ = "UPDATE HRPASDEP SET PD_USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE PD_USERID='" & Replace(OUserID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
SQLQ = "UPDATE HR_EMAIL SET EM_USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE EM_USERID='" & Replace(OUserID, "'", "''") & "'"
gdbAdoIhr001.Execute SQLQ
If glbLinamar Then
    SQLQ = "UPDATE LN_SECURE_ACCESS SET USERID='" & Replace(txtUSERID, "'", "''") & "' WHERE USERID='" & Replace(OUserID, "'", "''") & "'"
    gdbAdoIhr001.Execute SQLQ
End If
End Sub

Private Sub SECWRK()

Dim SQLQ, xField As String, x
Dim xQue
Dim rsSEC As New ADODB.Recordset
Dim rsFun As New ADODB.Recordset
Dim rsSECWrk As New ADODB.Recordset

SQLQ = "select * from HR_SECURE_BASIC where USERID='" & Replace(txtUSERID, "'", "''") & "'"
rsSEC.Open SQLQ, gdbAdoIhr001, adOpenStatic
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRSECWRK"
gdbAdoIhr001W.CommitTrans
rsSECWrk.Open "HRSECWRK", gdbAdoIhr001W, adOpenStatic, adLockPessimistic

Do Until rsSEC.EOF
    rsSECWrk.AddNew
    rsSECWrk("EMPNBR") = rsSEC("EMPNBR")
    rsSECWrk("USERID") = rsSEC("USERID")
    xQue = False
    For x = 1 To rsSECWrk.Fields.count - 1
        xField = rsSECWrk.Fields(x).name
        xQue = True
        If UCase(xField) = UCase("Basic_Inquiry") Then xQue = True
        If UCase(xField) = "PS_CHGDATE" Then xQue = False
        If xQue Then
            rsFun.Open "select * from hr_secure_ACCESS where userid='" & Replace(rsSEC("USERID"), "'", "''") & "' and [function]='" & xField & "'", gdbAdoIhr001, adOpenStatic
            If Not rsFun.EOF Then rsSECWrk(xField) = rsFun("Accessable")
            rsFun.Close
        End If
    Next
    rsSECWrk("WRKEMP") = glbUserID
    rsSECWrk.Update
    rsSEC.MoveNext
Loop
End Sub
Private Sub ResetAll()
Dim x%, starttimer As Single

For x% = 0 To 6
    chkMSecurity(x%).Value = 0
    chkSecurity(x%).Value = 0
Next x%

chkRCompT.Value = 0
End Sub

Private Sub txtPWord_Change()
If gsMultiLang = "YES" Then
    txtSecPassword.Text = EncryptPasswordMultiLang(txtPWord.Text)
Else
    txtSecPassword.Text = EncryptPassword(txtPWord.Text)
End If
End Sub

Private Sub txtPWord_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSecPassword_Change()
If gsMultiLang = "YES" Then
    txtPWord.Text = DecryptPasswordMultiLang(txtSecPassword.Text)
Else
    txtPWord.Text = DecryptPassword(txtSecPassword.Text)
End If
End Sub

Private Sub txtUSERID_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Dim VR
Call ChkCBoxChange
If ChangeCBox = True Then
        VR = MsgBox("Do you want to save changes?", MB_YESNO)
        If VR = IDYES Then
            Me.cmdOK_Click
        ElseIf VR = IDNO Then
            Call Me.cmdCancel_Click
        End If
End If

If gsMultiLang = "YES" Then
    If txtPWord.Text <> DecryptPasswordMultiLang(Data1.Recordset("PassWord")) Then
        glbConfPass = txtPWord
        Load frmConfPass
        frmConfPass.Show vbModal
    End If
Else
    If txtPWord.Text <> DecryptPassword(Data1.Recordset("PassWord")) Then
        glbConfPass = txtPWord
        Load frmConfPass
        frmConfPass.Show vbModal
    End If
End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT PD_DEPT FROM HRPASDEP WHERE PD_USERID = '" & Replace(txtUSERID, "'", "''") & "'"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
   '     cmdOK.SetFocus
  '  Else
 '       cmdModify.SetFocus
'    End If
End If

End Sub
Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

Dim SQLE
Dim rsEMP As New ADODB.Recordset
Dim VR
On Error GoTo Tab1_Err

If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
   glbSecUSERID = txtUSERID
   lstDept.Clear
   If Not fglbNew Then
   
        SQLE = "SELECT PD_DEPT FROM HRPASDEP WHERE PD_USERID = '" & Replace(txtUSERID, "'", "''") & "'"
        rsEMP.Open SQLE, gdbAdoIhr001, adOpenStatic
       
        Do While Not rsEMP.EOF
            lstDept.AddItem (rsEMP("PD_DEPT"))
            rsEMP.MoveNext
        Loop

    Else
        lstDept.Clear
        lstDept.AddItem "Double Click"
   End If

   Call Display_Values
End If

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_OCC_HEALTH_SAFETY", "Add")
Call RollBack

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
UpdateRight = gSec_Upd_Quick_ESS
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
If fglbNew Then
    UpdateState = NewRecord
    TF = True
vbxTrueGrid.Enabled = False
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
vbxTrueGrid.Enabled = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call mod_UpdateMode(TF)

End Sub

Private Sub ChkCBoxChange()
Dim rsCB As New ADODB.Recordset
Dim x%, SQLQ
Dim xAccessable
SQLQ = "SELECT " & Field_SQL("FUNCTION") & ", ACCESSABLE "
SQLQ = SQLQ & " FROM HR_SECURE_ACCESS "
SQLQ = SQLQ & " WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND Maintainable=0"
rsCB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
ChangeCBox = False
Do Until rsCB.EOF
    If glbOracle Then
        If rsCB("ACCESSABLE") = 1 Then
            xAccessable = True
        Else
            xAccessable = False
        End If
    Else
        xAccessable = rsCB("ACCESSABLE")
    End If
    If UCase(rsCB("FUNCTION")) = UCase("Basic_Update") And xAccessable <> chkMSecurity(0) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Basic_Inquiry") And xAccessable <> chkSecurity(0) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Dependents_Update") And xAccessable <> chkMSecurity(1) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Dependents_Inquiry") And xAccessable <> chkSecurity(1) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Benefits_Update") And xAccessable <> chkMSecurity(2) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Benefits_Inquiry") And xAccessable <> chkSecurity(2) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Entitlements_Update") And xAccessable <> chkMSecurity(3) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Entitlements_Inquiry") And xAccessable <> chkSecurity(3) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Attendance_Update") And xAccessable <> chkMSecurity(4) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Attendance_Inquiry") And xAccessable <> chkSecurity(4) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Hrly_Entitlements_Update") And xAccessable <> chkMSecurity(5) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Hrly_Entitlements_Inquiry") And xAccessable <> chkSecurity(5) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Counselling_Update") And xAccessable <> chkMSecurity(6) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Counselling_Inquiry") And xAccessable <> chkSecurity(6) Then GoTo TheEnd
    
    'sam added 07/27/2006 Ticket # 11043
    If UCase(rsCB("FUNCTION")) = UCase("Archive_VacTimeoff_Update") And xAccessable <> chkMSecurity(7) Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Archive_VacTimeoff_Inquiry") And xAccessable <> chkSecurity(7) Then GoTo TheEnd
    'ends
    
    If UCase(rsCB("FUNCTION")) = UCase("Show_SIN_SSN") And xAccessable <> chkEESIN Then GoTo TheEnd
    If UCase(rsCB("FUNCTION")) = UCase("Report_Compensatory_Time") And xAccessable <> chkRCompT Then GoTo TheEnd
    rsCB.MoveNext
Loop
Exit Sub
TheEnd:
ChangeCBox = True
End Sub
