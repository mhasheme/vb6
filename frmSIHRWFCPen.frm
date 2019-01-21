VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSIHRWFCPen 
   Caption         =   "Pension System Security"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   10140
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Caption         =   "C.A.R.S. Administration Report"
      Height          =   8715
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   11355
      Begin VB.CommandButton cmdRemoveAll 
         Appearance      =   0  'Flat
         Caption         =   "&Remove All"
         Height          =   330
         Left            =   2640
         TabIndex        =   53
         Tag             =   "Grant All Basic"
         Top             =   8400
         Width           =   1200
      End
      Begin VB.CommandButton cmdGrantInqu 
         Appearance      =   0  'Flat
         Caption         =   "Grant All &Inquire"
         Height          =   330
         Left            =   3900
         TabIndex        =   52
         Tag             =   "Grant All Basic"
         Top             =   8400
         Width           =   1320
      End
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         Height          =   330
         Left            =   5400
         TabIndex        =   10
         Top             =   8400
         Width           =   1305
      End
      Begin Threed.SSCheck chkLSecurity 
         Height          =   225
         Index           =   11
         Left            =   375
         TabIndex        =   11
         Top             =   2970
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Annual Maximums"
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
         Left            =   375
         TabIndex        =   12
         Top             =   3840
         Width           =   3525
         _Version        =   65536
         _ExtentX        =   6218
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Hourly/Salaried Benefit Statement"
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
         Left            =   375
         TabIndex        =   13
         Top             =   4260
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Rates Master"
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
         Left            =   4095
         TabIndex        =   14
         Top             =   2970
         Width           =   2205
         _Version        =   65536
         _ExtentX        =   3889
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PA Details"
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
         Left            =   4095
         TabIndex        =   15
         Top             =   3180
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PA Master"
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
         Left            =   4095
         TabIndex        =   16
         Top             =   3390
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PA To Payroll"
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
         Left            =   4095
         TabIndex        =   17
         Top             =   3600
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PA Variance"
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
         Left            =   7440
         TabIndex        =   18
         Top             =   3390
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Master"
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
         Left            =   7440
         TabIndex        =   19
         Top             =   3600
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Beneficiary"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   7680
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "ADP Year End ECR"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   7320
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
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
         Left            =   7440
         TabIndex        =   25
         Top             =   4020
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7056
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Towers Watson Reports"
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
         Left            =   375
         TabIndex        =   26
         Top             =   4800
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "ADP Paytots"
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
         Index           =   42
         Left            =   4560
         TabIndex        =   27
         Top             =   8160
         Width           =   945
         _Version        =   65536
         _ExtentX        =   1667
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Sunlife"
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
         Height          =   225
         Index           =   5
         Left            =   375
         TabIndex        =   29
         Top             =   1065
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   4
         Left            =   375
         TabIndex        =   30
         Top             =   840
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   2
         Left            =   375
         TabIndex        =   31
         Top             =   6540
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         Height          =   225
         Index           =   1
         Left            =   375
         TabIndex        =   32
         Top             =   6345
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         Bindings        =   "frmSIHRWFCPen.frx":0000
         Height          =   225
         Index           =   1
         Left            =   1350
         TabIndex        =   33
         Top             =   6345
         Width           =   3765
         _Version        =   65536
         _ExtentX        =   6641
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Credited Service Rules"
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
         Bindings        =   "frmSIHRWFCPen.frx":000B
         Height          =   225
         Index           =   5
         Left            =   1350
         TabIndex        =   34
         Top             =   1065
         Width           =   4005
         _Version        =   65536
         _ExtentX        =   7064
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PA Master"
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
         Bindings        =   "frmSIHRWFCPen.frx":0016
         Height          =   225
         Index           =   4
         Left            =   1350
         TabIndex        =   35
         Top             =   840
         Width           =   1725
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "PA Details"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRWFCPen.frx":0021
         Height          =   225
         Index           =   2
         Left            =   1350
         TabIndex        =   36
         Top             =   6540
         Width           =   4665
         _Version        =   65536
         _ExtentX        =   8229
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Hourly/Salaried Benefit Statement Variable Setup"
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
         Bindings        =   "frmSIHRWFCPen.frx":002C
         Height          =   225
         Index           =   0
         Left            =   1350
         TabIndex        =   37
         Top             =   5880
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "Annual Maximum"
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
         Height          =   225
         Index           =   0
         Left            =   375
         TabIndex        =   38
         Top             =   5880
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         Height          =   225
         Index           =   10
         Left            =   375
         TabIndex        =   42
         Top             =   6960
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         Height          =   225
         Index           =   9
         Left            =   375
         TabIndex        =   43
         Top             =   6750
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         Height          =   225
         Index           =   8
         Left            =   5640
         TabIndex        =   44
         Top             =   1305
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   7
         Left            =   5640
         TabIndex        =   45
         Top             =   1065
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkMSecurity 
         Height          =   225
         Index           =   6
         Left            =   5640
         TabIndex        =   46
         Top             =   840
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRWFCPen.frx":0037
         Height          =   225
         Index           =   6
         Left            =   6615
         TabIndex        =   47
         Top             =   840
         Width           =   1605
         _Version        =   65536
         _ExtentX        =   2831
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Alerts"
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
         Bindings        =   "frmSIHRWFCPen.frx":0042
         Height          =   225
         Index           =   10
         Left            =   1350
         TabIndex        =   48
         Top             =   6960
         Width           =   2925
         _Version        =   65536
         _ExtentX        =   5159
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Type Matrix"
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
         Bindings        =   "frmSIHRWFCPen.frx":004D
         Height          =   225
         Index           =   9
         Left            =   1350
         TabIndex        =   49
         Top             =   6750
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Rates Master"
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
         Bindings        =   "frmSIHRWFCPen.frx":0058
         Height          =   225
         Index           =   7
         Left            =   6615
         TabIndex        =   50
         Top             =   1065
         Width           =   3225
         _Version        =   65536
         _ExtentX        =   5689
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Beneficiary"
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
         Bindings        =   "frmSIHRWFCPen.frx":0063
         Height          =   225
         Index           =   8
         Left            =   6615
         TabIndex        =   51
         Top             =   1305
         Width           =   2745
         _Version        =   65536
         _ExtentX        =   4842
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Master"
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
         Index           =   0
         Left            =   1680
         TabIndex        =   55
         Top             =   0
         Width           =   525
         _Version        =   65536
         _ExtentX        =   926
         _ExtentY        =   397
         _StockProps     =   78
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
         Height          =   225
         Index           =   11
         Left            =   375
         TabIndex        =   56
         Top             =   6120
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
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
         Bindings        =   "frmSIHRWFCPen.frx":006E
         Height          =   225
         Index           =   11
         Left            =   1350
         TabIndex        =   57
         Top             =   6120
         Width           =   3885
         _Version        =   65536
         _ExtentX        =   6853
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "CIA Interest Rate Matrix"
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
         Left            =   4095
         TabIndex        =   58
         Top             =   4800
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Government Forms"
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
         Index           =   21
         Left            =   7440
         TabIndex        =   59
         Top             =   2970
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Calculators - Estimate"
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
         Index           =   22
         Left            =   375
         TabIndex        =   63
         Top             =   3630
         Width           =   2205
         _Version        =   65536
         _ExtentX        =   3889
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Headcount Audit"
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
         Left            =   375
         TabIndex        =   64
         Top             =   4050
         Width           =   2205
         _Version        =   65536
         _ExtentX        =   3889
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Hourly PA Worksheet"
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
         Height          =   225
         Index           =   16
         Left            =   375
         TabIndex        =   65
         Top             =   1305
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRWFCPen.frx":0079
         Height          =   225
         Index           =   16
         Left            =   1350
         TabIndex        =   66
         Top             =   1305
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PAR Master"
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
         Height          =   225
         Index           =   17
         Left            =   375
         TabIndex        =   67
         Top             =   1530
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRWFCPen.frx":0084
         Height          =   225
         Index           =   17
         Left            =   1350
         TabIndex        =   68
         Top             =   1530
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PAR Update"
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
         Left            =   4095
         TabIndex        =   69
         Top             =   3810
         Width           =   2805
         _Version        =   65536
         _ExtentX        =   4948
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PAR Preparation Worksheet"
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
         Left            =   4095
         TabIndex        =   70
         Top             =   4020
         Width           =   2205
         _Version        =   65536
         _ExtentX        =   3889
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PAR Report"
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
         Left            =   375
         TabIndex        =   71
         Top             =   5040
         Width           =   1905
         _Version        =   65536
         _ExtentX        =   3360
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PSPA Load"
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
         Left            =   6645
         TabIndex        =   72
         Top             =   5760
         Width           =   2385
         _Version        =   65536
         _ExtentX        =   4207
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Archive"
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
         Left            =   6645
         TabIndex        =   73
         Top             =   6000
         Width           =   2025
         _Version        =   65536
         _ExtentX        =   3572
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Overview/Summary"
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
         Index           =   26
         Left            =   375
         TabIndex        =   77
         Top             =   3180
         Width           =   2685
         _Version        =   65536
         _ExtentX        =   4736
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Credited Service Report"
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
         Index           =   27
         Left            =   7440
         TabIndex        =   78
         Top             =   3810
         Width           =   2325
         _Version        =   65536
         _ExtentX        =   4101
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Sunlife"
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
      Begin Threed.SSCheck chkSuSecurity 
         Bindings        =   "frmSIHRWFCPen.frx":008F
         Height          =   225
         Index           =   6
         Left            =   8280
         TabIndex        =   79
         Top             =   840
         Width           =   4485
         _Version        =   65536
         _ExtentX        =   7911
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Alerts Mark/Delete"
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
         Index           =   28
         Left            =   4095
         TabIndex        =   80
         Top             =   4260
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   3466
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Alerts"
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
         Left            =   4095
         TabIndex        =   81
         Top             =   5040
         Width           =   2265
         _Version        =   65536
         _ExtentX        =   3995
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Load EE Contributions"
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
         Left            =   375
         TabIndex        =   82
         Top             =   3390
         Width           =   3285
         _Version        =   65536
         _ExtentX        =   5794
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Employee DB Contribution Report"
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
         Index           =   30
         Left            =   7440
         TabIndex        =   83
         Top             =   3180
         Width           =   3045
         _Version        =   65536
         _ExtentX        =   5371
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Pension Calculators - Actual"
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
         Left            =   4560
         TabIndex        =   85
         Top             =   7920
         Width           =   3585
         _Version        =   65536
         _ExtentX        =   6324
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Special Early Retirement Calculation"
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
         Left            =   1320
         TabIndex        =   86
         Top             =   8160
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Freeze DB Interest Rate"
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
         Left            =   4560
         TabIndex        =   87
         Top             =   7680
         Width           =   3825
         _Version        =   65536
         _ExtentX        =   6747
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Hourly Pension && PA Update At Yearend"
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
         Left            =   375
         TabIndex        =   89
         Top             =   2520
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Request an Estimate"
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
         Left            =   4095
         TabIndex        =   90
         Top             =   2280
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Termination Packages"
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
         Left            =   4095
         TabIndex        =   91
         Top             =   2040
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Retirement Packages"
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
         Left            =   375
         TabIndex        =   92
         Top             =   2040
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "PAR T10"
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
         Left            =   375
         TabIndex        =   93
         Top             =   2280
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Condolence Packages"
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
         Height          =   225
         Index           =   3
         Left            =   5640
         TabIndex        =   94
         Top             =   1530
         Width           =   435
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Font3D          =   3
      End
      Begin Threed.SSCheck chkSecurity 
         Bindings        =   "frmSIHRWFCPen.frx":009A
         Height          =   225
         Index           =   3
         Left            =   6615
         TabIndex        =   95
         Top             =   1530
         Width           =   3345
         _Version        =   65536
         _ExtentX        =   5900
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Employee DB Contribution"
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
         Left            =   4095
         TabIndex        =   96
         Top             =   2520
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Small Payout Packages"
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
         Left            =   1320
         TabIndex        =   97
         Top             =   7920
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "DBSERP Calculation"
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
         Left            =   5640
         TabIndex        =   28
         Top             =   8160
         Width           =   2985
         _Version        =   65536
         _ExtentX        =   5265
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Update Hourly DC Contribution"
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
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Forms/Packages"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   88
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year End"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   84
         Top             =   7680
         Width           =   660
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   6600
         TabIndex        =   76
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   5640
         TabIndex        =   75
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Archive Pension Data"
         Height          =   375
         Left            =   6360
         TabIndex        =   74
         Top             =   5400
         Width           =   1875
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1320
         TabIndex        =   62
         Top             =   5640
         Width           =   600
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   61
         Top             =   5640
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Setup"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   5400
         Width           =   780
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Overview/Summary"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintenance"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   90
         TabIndex        =   41
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maintain"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   40
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Inquire"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1380
         TabIndex        =   39
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Reports"
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label Label8 
         Caption         =   "Import"
         Height          =   375
         Left            =   90
         TabIndex        =   23
         Top             =   4560
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Initial Data Load"
         Height          =   255
         Left            =   90
         TabIndex        =   22
         Top             =   7320
         Width           =   1275
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
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
         TabIndex        =   2
         Top             =   125
         Width           =   630
      End
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
         TabIndex        =   1
         Top             =   120
         Width           =   630
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   9480
      Width           =   11715
      _Version        =   65536
      _ExtentX        =   20664
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
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1140
         TabIndex        =   8
         Tag             =   "Edit the information "
         Top             =   0
         Width           =   765
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   315
         TabIndex        =   7
         Tag             =   "Close and exit this screen"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2010
         TabIndex        =   6
         Tag             =   "Save the changes made"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2835
         TabIndex        =   5
         Tag             =   "Cancel the changes made"
         Top             =   0
         Width           =   795
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
End
Attribute VB_Name = "frmSIHRWFCPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer

Private Sub chkMSecurity_Click(Index As Integer, Value As Integer)
If chkMSecurity(Index).Value = True Then
    chkSecurity(Index).Value = True
End If
End Sub

Private Sub chkSecurity_Click(Index As Integer, Value As Integer)
    If chkSecurity(Index).Value = False Then
        chkMSecurity(Index).Value = False
    End If
End Sub

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

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGrantAll_Click()
Dim X%

For X% = 0 To 17
    If X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
    chkMSecurity(X%).Value = 1
    End If
Next X%
For X% = 0 To 17
    If X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
    chkSecurity(X%).Value = 1
    End If
Next X%

chkLSecurity(0).Value = 1
For X% = 11 To 30 '25
    chkLSecurity(X%).Value = 1
Next X%
For X% = 40 To 59 '58
    chkLSecurity(X%).Value = 1
Next X%
'chkLSecurity(50).Value = 1

chkSuSecurity(6).Value = 1
End Sub

Private Sub cmdGrantInqu_Click()
Dim X%

For X% = 0 To 17
    If X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
    chkSecurity(X%).Value = 1
    End If
Next X%

For X% = 11 To 30 '25
    chkLSecurity(X%).Value = 1
Next X%
'For X% = 40 To 42
'    chkLSecurity(X%).Value = 1
'Next X%
'chkLSecurity(50).Value = 1
End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

If Not gSec_Upd_Security Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

chkMSecurity(0).SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "PensionSecurity", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdOK_Click()
Dim X%
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
    '????Ticket #24808 - User's based on this Template does not need their Profile to be updated as we are now retrieving Template profile for the users
    'Call procedure to Update all users with this template.
    'Call Update_Users_withthis_Template(glbSecUSERID)
End If

fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "Pension Security", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdRemoveAll_Click()
Dim X%

For X% = 0 To 17
    'If x% = 3 Or x% = 12 Or x% = 13 Or x% = 14 Or x% = 15 Then
    If X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
    chkMSecurity(X%).Value = 0
    End If
Next X%
For X% = 0 To 17
    If X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
        chkSecurity(X%).Value = 0
    End If
Next X%

chkLSecurity(0).Value = 0
For X% = 11 To 30 '25
    chkLSecurity(X%).Value = 0
Next X%
For X% = 40 To 59 '58
    chkLSecurity(X%).Value = 0
Next X%
'chkLSecurity(50).Value = 0

chkSuSecurity(6).Value = 0
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
Dim xTemplate  As String

glbOnTop = Me.name
Screen.MousePointer = HOURGLASS

lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName
frmSIHRWFCPen.Show

Me.Caption = ("Pension System Security - ") & lblEEName

Data1.ConnectionString = glbAdoIHRDB

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    Data1.RecordSource = "select * from HR_SECURE_ACCESS where USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],4)='WFC_'"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    Data1.RecordSource = "select * from HR_SECURE_ACCESS where USERID='" & Replace(xTemplate, "'", "''") & "' AND LEFT([FUNCTION],4)='WFC_'"
End If
Data1.Refresh

Call Display_Values

Call ST_UPD_MODE(False)

'Ticket #20585 - Enable/Disable Grant All and Grant All Users buttons based on the type of user
xTemplate = Get_Template(glbSecUSERID)
If xTemplate = "" Then
    'User without Template - Grant All Users will update all users with no template
    'cmdGrantAll.Enabled = True
    'cmdGrantUsr.Enabled = True  'will only update users without Template.
Else
    'User with Template or Template type of User - Do not Grant All Users
    'cmdGrantUsr.Enabled = False
    
    'Template or User based on a Template
    If xTemplate <> "TEMPLATE" Then
        'Do not Grant All for Users based on a Template
        cmdGrantAll.Enabled = False
        cmdGrantInqu.Enabled = False
        cmdRemoveAll.Enabled = False
        cmdModify.Enabled = False
    Else
        'User is Template
        cmdGrantAll.Enabled = True
        cmdGrantInqu.Enabled = True
        cmdRemoveAll.Enabled = True
    End If
End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub Display_Values()
Dim rsSR As New ADODB.Recordset
Dim X%, SQLQ

Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbSecUSERID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],7)='WFCPEN_'"
Else
    '????Ticket #24808 -  Retrieve template's security profile
    SQLQ = "select * from HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "' AND LEFT([FUNCTION],7)='WFCPEN_'"
End If
rsSR.Open SQLQ, gdbAdoIhr001, adOpenStatic

Call ResetAll

Do Until rsSR.EOF
    'Overview/Summary
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_OverviewSummary") Then chkLSecurity(0) = rsSR("ACCESSABLE")
    'Maintain
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Annual_Maximum_Upt") Then chkMSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Credited_Service_Rules_Upt") Then chkMSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Benefit_Variable_Setup_Upt") Then chkMSecurity(2) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Hourly_PA_Update_Yearend_Upt") Then chkMSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_EmpDB_Contribu_Upt") Then chkMSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Details_Upt") Then chkMSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Master_Upt") Then chkMSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Alerts_Upt") Then chkMSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Beneficiary_Upt") Then chkMSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Master_Upt") Then chkMSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Rates_Master_Upt") Then chkMSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Type_Matrix_Upt") Then chkMSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_CIA_Rate_Matrix_Upt") Then chkMSecurity(11) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Request_Estimate_Upt") Then chkMSecurity(12) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Retire_Process_Upt") Then chkMSecurity(13) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Term_Process_Upt") Then chkMSecurity(14) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_PART10_Upt") Then chkMSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PARS_DETAILS_Upt") Then chkMSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PARS_UPDATE_Upt") Then chkMSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Alerts_Del") Then chkSuSecurity(6) = rsSR("ACCESSABLE")
    
    'Inquire
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Annual_Maximum_Inq") Then chkSecurity(0) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Credited_Service_Rules_Inq") Then chkSecurity(1) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Benefit_Variable_Setup_Inq") Then chkSecurity(2) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Hourly_PA_Update_Yearend_Inq") Then chkSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_EmpDB_Contribu_Inq") Then chkSecurity(3) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Details_Inq") Then chkSecurity(4) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Master_Inq") Then chkSecurity(5) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Alerts_Inq") Then chkSecurity(6) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Beneficiary_Inq") Then chkSecurity(7) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Master_Inq") Then chkSecurity(8) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Rates_Master_Inq") Then chkSecurity(9) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Type_Matrix_Inq") Then chkSecurity(10) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_CIA_Rate_Matrix_Inq") Then chkSecurity(11) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Request_Estimate_Inq") Then chkSecurity(12) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Retire_Process_Inq") Then chkSecurity(13) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Term_Process_Inq") Then chkSecurity(14) = rsSR("ACCESSABLE")
    'If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_PART10_Inq") Then chkSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PARS_DETAILS_Inq") Then chkSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PARS_UPDATE_Inq") Then chkSecurity(17) = rsSR("ACCESSABLE")
    
    'Form/Package
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Request_Estimate") Then chkLSecurity(52) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Retire_Process") Then chkLSecurity(53) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Term_Process") Then chkLSecurity(54) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_PART10") Then chkLSecurity(55) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_Condolence") Then chkLSecurity(56) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FP_SmallPayout") Then chkLSecurity(57) = rsSR("ACCESSABLE")

    'report
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Annual_Maximum_Rpt") Then chkLSecurity(11) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Hourly_Benefit_Statement_Rpt") Then chkLSecurity(12) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Rates_Master_Rpt") Then chkLSecurity(13) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Details_Rpt") Then chkLSecurity(14) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Master_Rpt") Then chkLSecurity(15) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_To_Payroll_Rpt") Then chkLSecurity(16) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PA_Variance_Rpt") Then chkLSecurity(17) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Master_Rpt") Then chkLSecurity(18) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Beneficiary_Rpt") Then chkLSecurity(19) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_TP_Report_Rpt") Then chkLSecurity(20) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PensionCalculations_Rpt") Then chkLSecurity(21) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_HeadcountAudit_Rpt") Then chkLSecurity(22) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_HourlyPAWorksheet_Rpt") Then chkLSecurity(23) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PARs_Prep_Rpt") Then chkLSecurity(24) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PAR_Rpt") Then chkLSecurity(25) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_CreditSvr_Rpt") Then chkLSecurity(26) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Sunlife_Rpt") Then chkLSecurity(27) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Pension_Alerts_Rpt") Then chkLSecurity(28) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_EmpDB_Contribu_Rpt") Then chkLSecurity(29) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PensionCalculations_RptACT") Then chkLSecurity(30) = rsSR("ACCESSABLE")
    
    'Import
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_ADP_Paytots_Imp") Then chkLSecurity(41) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Government_Forms_Imp") Then chkLSecurity(43) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_PSPA_Load_Imp") Then chkLSecurity(44) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_EE_CONTRIBUTION") Then chkLSecurity(47) = rsSR("ACCESSABLE")
    
    'Archive Pension
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_ARCHIVE") Then chkLSecurity(45) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_ARC_OVERVIEW") Then chkLSecurity(46) = rsSR("ACCESSABLE")
    
    'IDL
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_IDL") Then chkLSecurity(50) = rsSR("ACCESSABLE")

    'Year End
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_ADP_Year_End_ECR_Imp") Then chkLSecurity(40) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Sunlife_Imp") Then chkLSecurity(42) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_FreezeDB") Then chkLSecurity(48) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_SpecialEarlyRet") Then chkLSecurity(49) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Hourly_PA_Update_Yearend") Then chkLSecurity(51) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_DBSERP_Calc") Then chkLSecurity(58) = rsSR("ACCESSABLE")
    If UCase(rsSR("FUNCTION")) = UCase("WFCPEN_Upt_Hourly_DC") Then chkLSecurity(59) = rsSR("ACCESSABLE") 'Ticket #28118 Franks 02/02/2016
    
    rsSR.MoveNext
Loop

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

Private Sub UpdSecAccess()
Dim SQLQ

SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "' AND LEFT([FUNCTION],7)='WFCPEN_'"
gdbAdoIhr001.Execute SQLQ

Call AddSecAccess

End Sub

Private Sub AddSecAccess()
Dim SQLQ, sqlI

sqlI = "INSERT INTO HR_SECURE_ACCESS(COMPNO,USERID,[FUNCTION],ACCESSABLE) "
sqlI = sqlI & " VALUES('001','" & Replace(Trim(lblUSERID), "'", "''") & "',"

'Overview/Summary
SQLQ = sqlI & "'WFCPEN_OverviewSummary'," & IIf(chkLSecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'Maintain
SQLQ = sqlI & "'WFCPEN_Annual_Maximum_Upt'," & IIf(chkMSecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Credited_Service_Rules_Upt'," & IIf(chkMSecurity(1), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Benefit_Variable_Setup_Upt'," & IIf(chkMSecurity(2), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_Hourly_PA_Update_Yearend_Upt'," & IIf(chkMSecurity(3), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_EmpDB_Contribu_Upt'," & IIf(chkMSecurity(3), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Details_Upt'," & IIf(chkMSecurity(4), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Master_Upt'," & IIf(chkMSecurity(5), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Alerts_Upt'," & IIf(chkMSecurity(6), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Beneficiary_Upt'," & IIf(chkMSecurity(7), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Master_Upt'," & IIf(chkMSecurity(8), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Rates_Master_Upt'," & IIf(chkMSecurity(9), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Type_Matrix_Upt'," & IIf(chkMSecurity(10), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_CIA_Rate_Matrix_Upt'," & IIf(chkMSecurity(11), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_Request_Estimate_Upt'," & IIf(chkMSecurity(12), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_Retire_Process_Upt'," & IIf(chkMSecurity(13), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_Term_Process_Upt'," & IIf(chkMSecurity(14), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_PART10_Upt'," & IIf(chkMSecurity(15), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PARS_DETAILS_Upt'," & IIf(chkMSecurity(16), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PARS_UPDATE_Upt'," & IIf(chkMSecurity(17), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Alerts_Del'," & IIf(chkSuSecurity(6), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Inquire
SQLQ = sqlI & "'WFCPEN_Annual_Maximum_Inq'," & IIf(chkSecurity(0), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Credited_Service_Rules_Inq'," & IIf(chkSecurity(1), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Benefit_Variable_Setup_Inq'," & IIf(chkSecurity(2), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_Hourly_PA_Update_Yearend_Inq'," & IIf(chkSecurity(3), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_EmpDB_Contribu_Inq'," & IIf(chkSecurity(3), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Details_Inq'," & IIf(chkSecurity(4), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Master_Inq'," & IIf(chkSecurity(5), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Alerts_Inq'," & IIf(chkSecurity(6), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Beneficiary_Inq'," & IIf(chkSecurity(7), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Master_Inq'," & IIf(chkSecurity(8), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Rates_Master_Inq'," & IIf(chkSecurity(9), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Type_Matrix_Inq'," & IIf(chkSecurity(10), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_CIA_Rate_Matrix_Inq'," & IIf(chkSecurity(11), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_Request_Estimate_Inq'," & IIf(chkSecurity(12), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_Retire_Process_Inq'," & IIf(chkSecurity(13), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_Term_Process_Inq'," & IIf(chkSecurity(14), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
'SQLQ = sqlI & "'WFCPEN_FP_PART10_Inq'," & IIf(chkSecurity(15), 1, 0) & ")"
'gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PARS_DETAILS_Inq'," & IIf(chkSecurity(16), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PARS_UPDATE_Inq'," & IIf(chkSecurity(17), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'report
SQLQ = sqlI & "'WFCPEN_Annual_Maximum_Rpt'," & IIf(chkLSecurity(11), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Hourly_Benefit_Statement_Rpt'," & IIf(chkLSecurity(12), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Rates_Master_Rpt'," & IIf(chkLSecurity(13), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Details_Rpt'," & IIf(chkLSecurity(14), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Master_Rpt'," & IIf(chkLSecurity(15), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_To_Payroll_Rpt'," & IIf(chkLSecurity(16), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PA_Variance_Rpt'," & IIf(chkLSecurity(17), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Master_Rpt'," & IIf(chkLSecurity(18), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Beneficiary_Rpt'," & IIf(chkLSecurity(19), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_TP_Report_Rpt'," & IIf(chkLSecurity(20), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PensionCalculations_Rpt'," & IIf(chkLSecurity(21), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_HeadcountAudit_Rpt'," & IIf(chkLSecurity(22), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_HourlyPAWorksheet_Rpt'," & IIf(chkLSecurity(23), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PARs_Prep_Rpt'," & IIf(chkLSecurity(24), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PAR_Rpt'," & IIf(chkLSecurity(25), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_CreditSvr_Rpt'," & IIf(chkLSecurity(26), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Sunlife_Rpt'," & IIf(chkLSecurity(27), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Pension_Alerts_Rpt'," & IIf(chkLSecurity(28), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_EmpDB_Contribu_Rpt'," & IIf(chkLSecurity(29), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PensionCalculations_RptACT'," & IIf(chkLSecurity(30), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Import
SQLQ = sqlI & "'WFCPEN_ADP_Paytots_Imp'," & IIf(chkLSecurity(41), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Government_Forms_Imp'," & IIf(chkLSecurity(43), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_PSPA_Load_Imp'," & IIf(chkLSecurity(44), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_EE_CONTRIBUTION'," & IIf(chkLSecurity(47), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Archive Pension
SQLQ = sqlI & "'WFCPEN_ARCHIVE'," & IIf(chkLSecurity(45), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_ARC_OVERVIEW'," & IIf(chkLSecurity(46), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'IDL
SQLQ = sqlI & "'WFCPEN_IDL'," & IIf(chkLSecurity(50), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

'Year End
SQLQ = sqlI & "'WFCPEN_ADP_Year_End_ECR_Imp'," & IIf(chkLSecurity(40), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Sunlife_Imp'," & IIf(chkLSecurity(42), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_FreezeDB'," & IIf(chkLSecurity(48), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_SpecialEarlyRet'," & IIf(chkLSecurity(49), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Hourly_PA_Update_Yearend'," & IIf(chkLSecurity(51), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_DBSERP_Calc'," & IIf(chkLSecurity(58), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_Upt_Hourly_DC'," & IIf(chkLSecurity(59), 1, 0) & ")" 'Ticket #28118 Franks 02/02/2016
gdbAdoIhr001.Execute SQLQ

'Form/Package
SQLQ = sqlI & "'WFCPEN_FP_Request_Estimate'," & IIf(chkLSecurity(52), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_FP_Retire_Process'," & IIf(chkLSecurity(53), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_FP_Term_Process'," & IIf(chkLSecurity(54), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_FP_PART10'," & IIf(chkLSecurity(55), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_FP_Condolence'," & IIf(chkLSecurity(56), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ
SQLQ = sqlI & "'WFCPEN_FP_SmallPayout'," & IIf(chkLSecurity(57), 1, 0) & ")"
gdbAdoIhr001.Execute SQLQ

End Sub

Private Sub ResetAll()
Dim X%

For X% = 0 To 17
    If X% = 3 Or X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
    chkMSecurity(X%).Value = 0
    End If
Next X%
For X% = 0 To 17
    If X% = 3 Or X% = 12 Or X% = 13 Or X% = 14 Or X% = 15 Then
    Else
        chkSecurity(X%).Value = 0
    End If
Next X%

chkLSecurity(0).Value = 0
For X% = 11 To 26 '25
    chkLSecurity(X%).Value = 0
Next X%
For X% = 40 To 58
    chkLSecurity(X%).Value = 0
Next X%
'chkLSecurity(50).Value = 0
chkSuSecurity(6).Value = 0

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
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "CUSTOMFEATURE_PEN")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

