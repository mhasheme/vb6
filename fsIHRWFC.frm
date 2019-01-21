VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSIHRWFC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Hourly Reports"
   ClientHeight    =   8940
   ClientLeft      =   465
   ClientTop       =   1410
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8940
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   11835
      _Version        =   65536
      _ExtentX        =   20876
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2835
         TabIndex        =   10
         Tag             =   "Cancel the changes made"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2010
         TabIndex        =   9
         Tag             =   "Save the changes made"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   315
         TabIndex        =   8
         Tag             =   "Close and exit this screen"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1140
         TabIndex        =   7
         Tag             =   "Edit the information "
         Top             =   0
         Width           =   765
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   4200
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11835
      _Version        =   65536
      _ExtentX        =   20876
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
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descr"
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
         Left            =   3030
         TabIndex        =   5
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   125
         Width           =   630
      End
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   660
      End
   End
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Caption         =   "C.A.R.S. Administration Report"
      Height          =   7755
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         Height          =   360
         Left            =   5160
         TabIndex        =   1
         Top             =   7200
         Width           =   1305
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   11
         Left            =   240
         TabIndex        =   11
         Top             =   570
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Attendance Code Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   12
         Left            =   240
         TabIndex        =   12
         Top             =   780
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Budget Absenteeism"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   13
         Left            =   240
         TabIndex        =   13
         Top             =   990
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Budget Headcount"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   14
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Division Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   15
         Left            =   240
         TabIndex        =   15
         Top             =   1410
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Employment Status Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   16
         Left            =   4680
         TabIndex        =   16
         Top             =   570
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Fiscal Month Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   17
         Left            =   4680
         TabIndex        =   17
         Top             =   780
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Termination Reason Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   18
         Left            =   4680
         TabIndex        =   18
         Top             =   990
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Total Monthly Hours Worked"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   19
         Left            =   4680
         TabIndex        =   19
         Top             =   1200
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Training Development Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   20
         Left            =   240
         TabIndex        =   20
         Top             =   1950
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "HR STATS Report"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   23
         Left            =   4680
         TabIndex        =   21
         Top             =   1950
         Width           =   5145
         _Version        =   65536
         _ExtentX        =   9075
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Status Change Report"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   24
         Left            =   240
         TabIndex        =   22
         Top             =   5040
         Width           =   3825
         _Version        =   65536
         _ExtentX        =   6747
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Maintain/Inquire Requisition"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   25
         Left            =   4680
         TabIndex        =   23
         Top             =   5040
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Requisition Approval/Decline "
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   29
         Left            =   240
         TabIndex        =   27
         Top             =   3600
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Band Security (A -> E)"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   38
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Import Other Earnings"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   40
         Left            =   240
         TabIndex        =   32
         Top             =   4080
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Database Setup"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   41
         Left            =   4680
         TabIndex        =   33
         Top             =   2880
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Import Dependents"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   43
         Left            =   4680
         TabIndex        =   34
         Top             =   4080
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "CoreSource Sub Group Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   44
         Left            =   4680
         TabIndex        =   36
         Top             =   4320
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Weekly Amount Matrix"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   45
         Left            =   4680
         TabIndex        =   37
         Top             =   4560
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Payroll Email Addresses"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   46
         Left            =   240
         TabIndex        =   38
         Top             =   2160
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Manulife Audit Report"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   47
         Left            =   4680
         TabIndex        =   39
         Top             =   2160
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "CoreSource Audit Report"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   48
         Left            =   4680
         TabIndex        =   41
         Top             =   3630
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Unlock Smoker Status"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   49
         Left            =   240
         TabIndex        =   42
         Top             =   4560
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Benefit Account Setup"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   50
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   3795
         _Version        =   65536
         _ExtentX        =   6694
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Woodbridge Entitlements Report"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   51
         Left            =   240
         TabIndex        =   45
         Top             =   3120
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Internal Phones and Email Addresses"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   52
         Left            =   270
         TabIndex        =   46
         Top             =   5490
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Currency Exchange Rate"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   53
         Left            =   270
         TabIndex        =   47
         Top             =   5700
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Company Incentive Factors"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   54
         Left            =   270
         TabIndex        =   48
         Top             =   5910
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Create Incentive Plan Spreadsheet"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   55
         Left            =   270
         TabIndex        =   49
         Top             =   6120
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Import Incentive Plan Spreadsheet"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   56
         Left            =   4680
         TabIndex        =   51
         Top             =   5520
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Update info:HR Other Earnings"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   57
         Left            =   4680
         TabIndex        =   52
         Top             =   5730
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Prepare Payroll Transaction File"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   58
         Left            =   4680
         TabIndex        =   53
         Top             =   5940
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Print Incentive Plan Spreadsheet"
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
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   59
         Left            =   4680
         TabIndex        =   54
         Top             =   6150
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Print Employee Incentive Letter"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Incentive Plan"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   5280
         Width           =   1245
      End
      Begin VB.Label Label3 
         Caption         =   "Manulife"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4320
         Width           =   1875
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Others"
         Height          =   195
         Left            =   4680
         TabIndex        =   40
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label Label10 
         Caption         =   "CoreSource"
         Height          =   255
         Left            =   4680
         TabIndex        =   35
         Top             =   3840
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "BIC Interface"
         Height          =   255
         Left            =   90
         TabIndex        =   31
         Top             =   3840
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "WFC Import"
         Height          =   375
         Left            =   90
         TabIndex        =   29
         Top             =   2640
         Width           =   1875
      End
      Begin VB.Label Label8 
         Caption         =   "Salary"
         Height          =   375
         Left            =   90
         TabIndex        =   28
         Top             =   3360
         Width           =   1875
      End
      Begin VB.Label Label6 
         Caption         =   "Position Posting"
         Height          =   375
         Left            =   90
         TabIndex        =   26
         Top             =   4800
         Width           =   1875
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "HR Stats"
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "Reports"
         Height          =   375
         Left            =   90
         TabIndex        =   24
         Top             =   1680
         Width           =   1875
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_Return 
         Caption         =   "&Return to Security"
      End
   End
End
Attribute VB_Name = "frmSIHRWFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer

Private Sub cmdCancel_Click()

On Error GoTo Can_Err

Call Display_Values
Call ST_UPD_MODE(False)  ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdCancel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmdClose_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdGrantAll_Click()
Dim x%
'For x% = 1 To 38
'For x% = 9 To 20
For x% = 11 To 20
    chkLSecurity(x%).Value = 1
Next x%

For x% = 23 To 25 '28
    chkLSecurity(x%).Value = 1
Next x%
'29 Band Security (A -> E): Default to false
'For x% = 30 To 48 '47 '45 '44 '43 '39 '36
'    chkLSecurity(x%).Value = 1
'Next x%

For x% = 38 To 51 '49 '48 '47 '45 '44 '43 '39 '36
    If x% = 39 Or x% = 42 Then
    Else
    chkLSecurity(x%).Value = 1
    End If
Next x%

For x% = 52 To 59 'Ticket #29846 Franks 03/07/2017
    chkLSecurity(x%).Value = 1
Next x%

End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Security Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

'chkLSecurity(9).SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "HRJOBEVL", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub



Private Sub cmdOK_Click()
Dim x%
Dim xID
Dim xTemplate As String

On Error GoTo OK_Err

Call ST_UPD_MODE(False)

'Ticket #20585 - If Template then update users with this template as well.
'If User and with no template then update that user's profile.
'if User and with Template then do not update user's profile.
'Get the Template Name of this User ID
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "TEMPLATE" Then
    'Update all users with this template. After the changes are saved
ElseIf xTemplate = "" Then
    'User - User with no template - don't do anything let system update user's profile
ElseIf xTemplate <> "TEMPLATE" Then
    'User with template - do not allow to save these changes.
    MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbInformation, "Template based User Security Profile"
    
    'Redisplay the security settings
    Call Display_Values
End If

'Template or User only
If xTemplate = "TEMPLATE" Or xTemplate = "" Then
    Call UpdSecAccess
End If

If xTemplate = "TEMPLATE" Then
    'Call procedure to Update all users with this template.
    Call Update_Users_withthis_Template(glbSecUSERID)
End If

fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBEVL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRJOBEVL", "SELECT")


End Sub



Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
glbOnTop = Me.name
Screen.MousePointer = HOURGLASS
lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName
frmSIHRWFC.Show
Me.Caption = lStr("WFC Security - ") & lblEEName
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "select * from HR_SECURE_ACCESS where USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='WFC_'"
Data1.Refresh

Call Display_Values


Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmSIHRWFC = Nothing

End Sub

Private Sub mnu_File_Exit_Click()
    Call ApplicationEnd
End Sub

Private Sub mnu_F_PrintSetup_Click()
MDIMain.vbxCommonDlg.Action = 5

End Sub

Private Sub mnu_Return_Click()
   Unload Me
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

glbOHSEdit% = TF

fUPMode = TF    ' update mode
cmdOK.Enabled = TF
cmdModify.Enabled = FT
cmdCancel.Enabled = TF
cmdClose.Enabled = FT
frmDetail.Enabled = TF
End Sub

Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim x%, SQLQ
SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='WFC_'"
rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic
Call ResetAll
Do Until rsSR.EOF
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Export_Control_Data_to_Plant") Then chkLSecurity(1) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Import_Into_Consolidation_Database") Then chkLSecurity(2) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Mass_Change_Payroll_To_Employee") Then chkLSecurity(3) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Import_Control_Data") Then chkLSecurity(4) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Export_For_Consolidation") Then chkLSecurity(5) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Year_End_Backup") Then chkLSecurity(6) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Create_Health_and_Safety_Files") Then chkLSecurity(7) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Export_Data_to_Head_Office") Then chkLSecurity(8) = rsSR("ACCESSABLE")
    
    'Ticket #22556 Franks 09/18/2012 - begin removed
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Transfer_Employee_Out") Then chkLSecurity(9) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Transfer_Employee_In") Then chkLSecurity(10) = rsSR("ACCESSABLE")
    'Ticket #22556 Franks 09/18/2012 - end
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Attendance_Code_Matrix") Then chkLSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Budget_Absenteeism") Then chkLSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Budget_Headcount") Then chkLSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Division_Matrix") Then chkLSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Employment_Status_Matrix") Then chkLSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Fiscal_Month_Matrix") Then chkLSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Termination_Reason_Matrix") Then chkLSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Total_Monthly_Hours_Worked") Then chkLSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Training_Development_Matrix") Then chkLSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_HR_STATS_Report") Then chkLSecurity(20) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Payroll_ID_Errors_Active") Then chkLSecurity(21) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Payroll_ID_Errors_Terminated") Then chkLSecurity(22) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Status_Change_Report") Then chkLSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_JobPosting_Requisition_MainInquire") Then chkLSecurity(24) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_JobPosting_Requisition_ApproveDecline") Then chkLSecurity(25) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Bonus_Intergration_Interface") Then chkLSecurity(26) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Bonus_Adminstration_Employee") Then chkLSecurity(27) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Bonus_Adminstration_Position") Then chkLSecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Band_Security") Then chkLSecurity(29) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Bonus_Adminstration_Report") Then chkLSecurity(31) = rsSR("ACCESSABLE")
    'Ticket #22556 Franks 09/18/2012 - begin - remove
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_CARS_Business") Then chkLSecurity(30) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_CARS_Functional") Then chkLSecurity(32) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_CARS_Agreed") Then chkLSecurity(33) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_CARS_RBlank") Then chkLSecurity(34) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_CARS_RDetailed") Then chkLSecurity(35) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_CARS_RAdmin") Then chkLSecurity(36) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_IMP_ConEducation") Then chkLSecurity(37) = rsSR("ACCESSABLE")
    'Ticket #22556 Franks 09/18/2012 - end
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IMP_OtherEarnings") Then chkLSecurity(38) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_IMP_SalaryGrid") Then chkLSecurity(39) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_BIC_DBSetup") Then chkLSecurity(40) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IMP_BeneDependents") Then chkLSecurity(41) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_IMP_Beneficiaries") Then chkLSecurity(42) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_NGS_SubGroupMatrix") Then chkLSecurity(43) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_PPRateMatrix") Then chkLSecurity(44) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_PayEmail") Then chkLSecurity(45) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_Manulife_Audit_Rpt") Then chkLSecurity(46) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_NGS_Audit_Rpt") Then chkLSecurity(47) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_UnlockSmokerStatus") Then chkLSecurity(48) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_BenefitAccSetup") Then chkLSecurity(49) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_EntitlementsReport") Then chkLSecurity(50) = rsSR("ACCESSABLE") 'Ticket #28254 Franks 04/06/2016
    If UCase(rsSR("FUNCTION")) = UCase("WFC_ImpInternalPhoneAddress") Then chkLSecurity(51) = rsSR("ACCESSABLE") 'Ticket #28254 Franks 04/06/2016
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_New_Hire_Hourly") Then chkLSecurity(28) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Termination_Hourly") Then chkLSecurity(29) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Alphabetical_List_of_Employees_Salaried") Then chkLSecurity(30) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Alpha_Salary_List_Salaried") Then chkLSecurity(31) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Seniority_by_Division/DOH_Salaried") Then chkLSecurity(32) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Seniority_by_DOH_Salaried") Then chkLSecurity(33) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Home_Address/Telephone_Salaried") Then chkLSecurity(34) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_New_Hire_Salaried") Then chkLSecurity(35) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Termination_Salaried") Then chkLSecurity(36) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Time_Card_Labels") Then chkLSecurity(37) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFC_Safety_Shoe_Report") Then chkLSecurity(38) = rsSR("ACCESSABLE")
    
    'Ticket #29846 Franks 03/07/2017 ----------------- begin
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPExchangeRate") Then chkLSecurity(52) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPIncentiveFactors") Then chkLSecurity(53) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPCreateSpreadsheet") Then chkLSecurity(54) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPImportSpreadsheet") Then chkLSecurity(55) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPUpdateEarnings") Then chkLSecurity(56) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPPreparePayrollFile") Then chkLSecurity(57) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPPrintSpreadsheet") Then chkLSecurity(58) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFC_IPPrintLetter") Then chkLSecurity(59) = rsSR("ACCESSABLE")
    'Ticket #29846 Franks 03/07/2017 ----------------- end
    rsSR.MoveNext
Loop

End Sub

Private Sub UpdSecAccess()
Dim SQLQ

SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='WFC_'"
gdbAdoIhr001.Execute SQLQ
Call AddSecAccess

End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI

sqlI = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID,[FUNCTION],ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(lblUSERID), "'", "''") & "',"

'SQLQ = sqlI & "'WFC_Export_Control_Data_to_Plant'," & IIf(chkLSecurity(1), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Import_Into_Consolidation_Database'," & IIf(chkLSecurity(2), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Mass_Change_Payroll_To_Employee'," & IIf(chkLSecurity(3), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Import_Control_Data'," & IIf(chkLSecurity(4), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Export_For_Consolidation'," & IIf(chkLSecurity(5), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Year_End_Backup'," & IIf(chkLSecurity(6), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Create_Health_and_Safety_Files'," & IIf(chkLSecurity(7), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Export_Data_to_Head_Office'," & IIf(chkLSecurity(8), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ

'SQLQ = sqlI & "'WFC_Transfer_Employee_Out'," & IIf(chkLSecurity(9), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Transfer_Employee_In'," & IIf(chkLSecurity(10), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Attendance_Code_Matrix'," & IIf(chkLSecurity(11), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Budget_Absenteeism'," & IIf(chkLSecurity(12), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Budget_Headcount'," & IIf(chkLSecurity(13), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Division_Matrix'," & IIf(chkLSecurity(14), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Employment_Status_Matrix'," & IIf(chkLSecurity(15), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Fiscal_Month_Matrix'," & IIf(chkLSecurity(16), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Termination_Reason_Matrix'," & IIf(chkLSecurity(17), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Total_Monthly_Hours_Worked'," & IIf(chkLSecurity(18), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Training_Development_Matrix'," & IIf(chkLSecurity(19), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_HR_STATS_Report'," & IIf(chkLSecurity(20), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Payroll_ID_Errors_Active'," & IIf(chkLSecurity(21), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Payroll_ID_Errors_Terminated'," & IIf(chkLSecurity(22), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Status_Change_Report'," & IIf(chkLSecurity(23), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_JobPosting_Requisition_MainInquire'," & IIf(chkLSecurity(24), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_JobPosting_Requisition_ApproveDecline'," & IIf(chkLSecurity(25), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Bonus_Intergration_Interface'," & IIf(chkLSecurity(26), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Bonus_Adminstration_Employee'," & IIf(chkLSecurity(27), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Bonus_Adminstration_Position'," & IIf(chkLSecurity(28), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_Band_Security'," & IIf(chkLSecurity(29), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_CARS_Business'," & IIf(chkLSecurity(30), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Bonus_Adminstration_Report'," & IIf(chkLSecurity(31), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_CARS_Functional'," & IIf(chkLSecurity(32), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_CARS_Agreed'," & IIf(chkLSecurity(33), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_CARS_RBlank'," & IIf(chkLSecurity(34), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_CARS_RDetailed'," & IIf(chkLSecurity(35), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_CARS_RAdmin'," & IIf(chkLSecurity(36), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ

'SQLQ = sqlI & "'WFC_IMP_ConEducation'," & IIf(chkLSecurity(37), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IMP_OtherEarnings'," & IIf(chkLSecurity(38), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_IMP_SalaryGrid'," & IIf(chkLSecurity(39), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_BIC_DBSetup'," & IIf(chkLSecurity(40), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IMP_BeneDependents'," & IIf(chkLSecurity(41), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_IMP_Beneficiaries'," & IIf(chkLSecurity(42), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_NGS_SubGroupMatrix'," & IIf(chkLSecurity(43), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_PPRateMatrix'," & IIf(chkLSecurity(44), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ 'Ticket #21220 Franks 11/28/2011
SQLQ = sqlI & "'WFC_PayEmail'," & IIf(chkLSecurity(45), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ 'Ticket #21295 Franks 12 / 8 / 2011
SQLQ = sqlI & "'WFC_Manulife_Audit_Rpt'," & IIf(chkLSecurity(46), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_NGS_Audit_Rpt'," & IIf(chkLSecurity(47), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_UnlockSmokerStatus'," & IIf(chkLSecurity(48), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_BenefitAccSetup'," & IIf(chkLSecurity(49), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_EntitlementsReport'," & IIf(chkLSecurity(50), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_ImpInternalPhoneAddress'," & IIf(chkLSecurity(51), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Ticket #29846 Franks 03/07/2017 ----------------- begin
SQLQ = sqlI & "'WFC_IPExchangeRate'," & IIf(chkLSecurity(52), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPIncentiveFactors'," & IIf(chkLSecurity(53), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPCreateSpreadsheet'," & IIf(chkLSecurity(54), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPImportSpreadsheet'," & IIf(chkLSecurity(55), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPUpdateEarnings'," & IIf(chkLSecurity(56), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPPreparePayrollFile'," & IIf(chkLSecurity(57), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPPrintSpreadsheet'," & IIf(chkLSecurity(58), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFC_IPPrintLetter'," & IIf(chkLSecurity(59), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'Ticket #29846 Franks 03/07/2017 ----------------- end

'SQLQ = sqlI & "'WFC_CAW_Seniority_by_DOH_Hourly'," & IIf(chkLSecurity(26), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Home_Address/Telephone_Hourly'," & IIf(chkLSecurity(27), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_New_Hire_Hourly'," & IIf(chkLSecurity(28), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Termination_Hourly'," & IIf(chkLSecurity(29), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Alphabetical_List_of_Employees_Salaried'," & IIf(chkLSecurity(30), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Alpha_Salary_List_Salaried'," & IIf(chkLSecurity(31), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Seniority_by_Division/DOH_Salaried'," & IIf(chkLSecurity(32), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Seniority_by_DOH_Salaried'," & IIf(chkLSecurity(33), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Home_Address/Telephone_Salaried'," & IIf(chkLSecurity(34), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_New_Hire_Salaried'," & IIf(chkLSecurity(35), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Termination_Salaried'," & IIf(chkLSecurity(36), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Time_Card_Labels'," & IIf(chkLSecurity(37), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFC_Safety_Shoe_Report'," & IIf(chkLSecurity(38), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ

End Sub

Private Sub ResetAll()
Dim x%

'For x% = 1 To 38
'For x% = 9 To 20
For x% = 11 To 20
    chkLSecurity(x%).Value = 0
Next x%
chkLSecurity(23).Value = 0

End Sub

Private Sub Update_Users_withthis_Template(xTemplate)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Retrieve all users associated with this changed Template
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & xTemplate & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            'Update each user with this changed Template
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "CUSTOMFEATURE")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

