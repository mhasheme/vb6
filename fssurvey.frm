VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSurveyData 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Survey Data"
   ClientHeight    =   11100
   ClientLeft      =   195
   ClientTop       =   1005
   ClientWidth     =   10950
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
   ScaleHeight     =   11100
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNAICS 
      Appearance      =   0  'Flat
      DataField       =   "EQ_NAICS"
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
      Left            =   7605
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   18
      Tag             =   "00-Enter NAICS code"
      Top             =   5460
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtProv 
      Appearance      =   0  'Flat
      DataField       =   "EQ_PROV"
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5460
      Width           =   375
   End
   Begin VB.TextBox OETYPE 
      Appearance      =   0  'Flat
      DataField       =   "EQ_TYPE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3990
      TabIndex        =   93
      Top             =   2775
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "fssurvey.frx":0000
      Left            =   2070
      List            =   "fssurvey.frx":0002
      TabIndex        =   1
      Tag             =   "Type: Active or Terminated "
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Frame frQ8 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   92
      Top             =   8880
      Width           =   1935
      Begin Threed.SSOption optQ8YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   42
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ8YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   41
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ7 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   91
      Top             =   8520
      Width           =   1935
      Begin Threed.SSOption optQ7YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ7YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   38
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ6 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   90
      Top             =   8160
      Width           =   1935
      Begin Threed.SSOption optQ6YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   36
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ6YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   35
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   89
      Top             =   7800
      Width           =   1935
      Begin Threed.SSOption optQ5YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ5YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   88
      Top             =   7440
      Width           =   1935
      Begin Threed.SSOption optQ4YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   30
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ4YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   87
      Top             =   7080
      Width           =   1935
      Begin Threed.SSOption optQ3YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ3YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   86
      Top             =   6720
      Width           =   1935
      Begin Threed.SSOption optQ2YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   24
         TabStop         =   0   'False
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ2YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   23
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.Frame frQ1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8160
      TabIndex        =   85
      Top             =   6360
      Width           =   1935
      Begin Threed.SSOption optQ1YesNo 
         Height          =   285
         Index           =   1
         Left            =   960
         TabIndex        =   21
         TabStop         =   0   'False
         Tag             =   "41-No"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "   No"
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optQ1YesNo 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Tag             =   "41-Yes"
         Top             =   0
         Width           =   795
         _Version        =   65536
         _ExtentX        =   1402
         _ExtentY        =   503
         _StockProps     =   78
         Caption         =   "  Yes "
         ForeColor       =   -2147483640
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
   End
   Begin VB.TextBox txtQ8 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q8A"
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
      Left            =   10200
      TabIndex        =   84
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   8880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ7 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q7A"
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
      Left            =   10200
      TabIndex        =   83
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   8520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ6 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q6A"
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
      Left            =   10200
      TabIndex        =   82
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ5 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q5A"
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
      Left            =   10200
      TabIndex        =   81
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   7800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ4 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q4A"
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
      Left            =   10200
      TabIndex        =   80
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ3 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q3A"
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
      Left            =   10200
      TabIndex        =   79
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ2 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q2A"
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
      Left            =   10200
      TabIndex        =   78
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   6720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQ1 
      Appearance      =   0  'Flat
      DataField       =   "EQ_Q1A"
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
      Left            =   10200
      TabIndex        =   77
      TabStop         =   0   'False
      Text            =   "N"
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtQues1 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6360
      Width           =   5955
   End
   Begin VB.TextBox txtQues2 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6720
      Width           =   5955
   End
   Begin VB.TextBox txtQues3 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7080
      Width           =   5955
   End
   Begin VB.TextBox txtQues4 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7440
      Width           =   5955
   End
   Begin VB.TextBox txtQues5 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7800
      Width           =   5955
   End
   Begin VB.TextBox txtQues6 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   8160
      Width           =   5955
   End
   Begin VB.TextBox txtQues7 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8520
      Width           =   5955
   End
   Begin VB.TextBox txtQues8 
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
      Left            =   2070
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8880
      Width           =   5955
   End
   Begin VB.TextBox txtGender 
      Appearance      =   0  'Flat
      DataField       =   "EQ_EESEX"
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
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   8
      Tag             =   "00-Enter M or F"
      Top             =   5130
      Width           =   375
   End
   Begin VB.TextBox txtNOC 
      Appearance      =   0  'Flat
      DataField       =   "EQ_NOGC"
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
      Left            =   7605
      TabIndex        =   15
      Top             =   4470
      Width           =   735
   End
   Begin INFOHR_Controls.DateLookup dlpTermDate 
      DataField       =   "EQ_DOT"
      Height          =   285
      Left            =   7290
      TabIndex        =   17
      Tag             =   "40-Enter Date of Termination"
      Top             =   5130
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDueDate 
      DataField       =   "EQ_DUEDATE"
      Height          =   285
      Left            =   7290
      TabIndex        =   16
      Tag             =   "40-Enter Due Date"
      Top             =   4800
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpNOGC 
      DataField       =   "EQ_NOGC"
      Height          =   285
      Left            =   6360
      TabIndex        =   45
      Tag             =   "00-Enter N.O.C."
      Top             =   9720
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   6
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EQ_REGION"
      Height          =   285
      Index           =   1
      Left            =   7290
      TabIndex        =   14
      Tag             =   "00-Enter Region / CMA"
      Top             =   4140
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
      MaxLength       =   7
   End
   Begin INFOHR_Controls.DateLookup dlpSurvDate 
      DataField       =   "EQ_SURVEY"
      Height          =   285
      Left            =   1755
      TabIndex        =   6
      Tag             =   "41-Enter Survey Date"
      Top             =   4470
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      DataField       =   "EQ_EEPT"
      Height          =   285
      Left            =   1755
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   4140
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
      MaxLength       =   7
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      DataField       =   "EQ_EMPNBR"
      Height          =   285
      Left            =   1755
      TabIndex        =   3
      Tag             =   "11-Enter Employee No."
      Top             =   3480
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpPlanNbr 
      DataField       =   "EQ_PLAN"
      Height          =   285
      Left            =   1755
      TabIndex        =   2
      Tag             =   "11-Enter Plan Number"
      Top             =   3150
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   10
      LookupType      =   7
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fssurvey.frx":0004
      Height          =   2325
      Left            =   0
      OleObjectBlob   =   "fssurvey.frx":0018
      TabIndex        =   0
      Top             =   210
      Width           =   10815
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5880
      Top             =   9840
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   65
      Top             =   10440
      Width           =   10950
      _Version        =   65536
      _ExtentX        =   19315
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
      Begin VB.CommandButton cmdRecal 
         Caption         =   "&Recalculate"
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   120
         Width           =   1275
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8760
         Top             =   165
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.TextBox txtDisability 
      Appearance      =   0  'Flat
      DataField       =   "EQ_DISAYN"
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
      Left            =   7605
      MaxLength       =   1
      TabIndex        =   13
      Tag             =   "00-Enter Y or N"
      Top             =   3810
      Width           =   375
   End
   Begin VB.TextBox txtVisibMinority 
      Appearance      =   0  'Flat
      DataField       =   "EQ_VMYN"
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
      Left            =   7605
      MaxLength       =   1
      TabIndex        =   12
      Tag             =   "00-Enter Y or N"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txtAborig 
      Appearance      =   0  'Flat
      DataField       =   "EQ_ABORYN"
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
      Left            =   7605
      MaxLength       =   1
      TabIndex        =   11
      Tag             =   "00-Enter Y or N"
      Top             =   3150
      Width           =   375
   End
   Begin VB.TextBox txtOrg 
      Appearance      =   0  'Flat
      DataField       =   "EQ_ORG"
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
      Left            =   4320
      TabIndex        =   57
      Top             =   9780
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      DataField       =   "EQ_FNAME"
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
      Left            =   3780
      TabIndex        =   56
      Top             =   9780
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSurname 
      Appearance      =   0  'Flat
      DataField       =   "EQ_SURNAME"
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
      Left            =   3240
      TabIndex        =   55
      Top             =   9780
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSurveyCompleted 
      Appearance      =   0  'Flat
      DataField       =   "EQ_SURCOMP"
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
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   7
      Tag             =   "01-Enter Y / N"
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox txtEmpEquity 
      Appearance      =   0  'Flat
      DataField       =   "EQ_EENO"
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
      Left            =   2070
      MaxLength       =   8
      TabIndex        =   4
      Tag             =   "11-Enter Emp. Equity No."
      Top             =   3810
      Width           =   1275
   End
   Begin VB.TextBox Updstats 
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
      Index           =   2
      Left            =   7605
      TabIndex        =   44
      Top             =   3810
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Updstats 
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
      Index           =   0
      Left            =   4920
      MaxLength       =   12
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   9780
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
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
      Height          =   315
      Index           =   1
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   9780
      Visible         =   0   'False
      Width           =   645
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      Height          =   285
      Left            =   2485
      TabIndex        =   9
      Tag             =   "31-Province - Code"
      Top             =   5460
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EQ_ORGT1"
      Height          =   285
      Index           =   6
      Left            =   1755
      TabIndex        =   10
      Tag             =   "00-Orgranization - Code"
      Top             =   5790
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ORGN"
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NAICS"
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
      Left            =   5235
      TabIndex        =   97
      Top             =   5505
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Organization 1"
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
      Left            =   60
      TabIndex        =   96
      Top             =   5835
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
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
      Left            =   60
      TabIndex        =   95
      Top             =   5505
      Width           =   630
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   94
      Top             =   2820
      Width           =   480
   End
   Begin VB.Label lblQues1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #1"
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
      Left            =   120
      TabIndex        =   76
      Top             =   6390
      Width           =   1215
   End
   Begin VB.Label lblQues2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #2"
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
      Left            =   120
      TabIndex        =   75
      Top             =   6750
      Width           =   1215
   End
   Begin VB.Label lblQues3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #3"
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
      Left            =   120
      TabIndex        =   74
      Top             =   7110
      Width           =   1215
   End
   Begin VB.Label lblQues4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #4"
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
      Left            =   120
      TabIndex        =   73
      Top             =   7470
      Width           =   1215
   End
   Begin VB.Label lblQues5 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #5"
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
      Left            =   120
      TabIndex        =   72
      Top             =   7830
      Width           =   1215
   End
   Begin VB.Label lblQues6 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #6"
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
      Left            =   120
      TabIndex        =   71
      Top             =   8190
      Width           =   1215
   End
   Begin VB.Label lblQues7 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #7"
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
      Left            =   120
      TabIndex        =   70
      Top             =   8550
      Width           =   1095
   End
   Begin VB.Label lblQues8 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Question #8"
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
      Left            =   120
      TabIndex        =   69
      Top             =   8910
      Width           =   1215
   End
   Begin VB.Label lblNOCDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      TabIndex        =   68
      Top             =   4545
      Width           =   2445
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   60
      TabIndex        =   66
      Top             =   4170
      Width           =   870
   End
   Begin VB.Label lblTermDate 
      Appearance      =   0  'Flat
      Caption         =   "Date of Termination"
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
      Left            =   5235
      TabIndex        =   64
      Top             =   5145
      Width           =   1815
   End
   Begin VB.Label lblDueDate 
      Appearance      =   0  'Flat
      Caption         =   "Due Date"
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
      Left            =   5235
      TabIndex        =   63
      Top             =   4815
      Width           =   975
   End
   Begin VB.Label lblNOGC 
      Appearance      =   0  'Flat
      Caption         =   "N.O.C."
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
      Left            =   5235
      TabIndex        =   62
      Top             =   4485
      Width           =   975
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "Region / CMA"
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
      Left            =   5235
      TabIndex        =   61
      Top             =   4185
      Width           =   1020
   End
   Begin VB.Label lblDisibility 
      Appearance      =   0  'Flat
      Caption         =   "Disability"
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
      Left            =   5235
      TabIndex        =   60
      Top             =   3825
      Width           =   975
   End
   Begin VB.Label lblVisib 
      Appearance      =   0  'Flat
      Caption         =   "Visible Minority"
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
      Left            =   5235
      TabIndex        =   59
      Top             =   3495
      Width           =   1575
   End
   Begin VB.Label lblAborig 
      Appearance      =   0  'Flat
      Caption         =   "Aboriginal"
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
      Left            =   5235
      TabIndex        =   58
      Top             =   3165
      Width           =   1095
   End
   Begin VB.Label lblGender 
      Appearance      =   0  'Flat
      Caption         =   "Gender"
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
      Left            =   60
      TabIndex        =   54
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label lblSurveyCompl 
      Appearance      =   0  'Flat
      Caption         =   "Survey Completed"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   53
      Top             =   4830
      Width           =   1695
   End
   Begin VB.Label lblSurveyDate 
      Appearance      =   0  'Flat
      Caption         =   "Survey Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   52
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label lblEmpEquity 
      Appearance      =   0  'Flat
      Caption         =   "Emp. Equity No."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   51
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label lblEmplNo 
      Appearance      =   0  'Flat
      Caption         =   "Employee No."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   50
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Label lblPlanNbr 
      Appearance      =   0  'Flat
      Caption         =   "Plan Number"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   49
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "EQ_COMPNO"
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
      Left            =   2820
      TabIndex        =   48
      Top             =   9810
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmSurveyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEEID&
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNewRec%
Dim fglbNew As Boolean
'Dim NOC_Snap As New ADODB.Recordset
'Dim PlanNo_Snap As New ADODB.Recordset
Dim FLNames_Snap As New ADODB.Recordset
Dim fndFName As String, fndLName As String
Dim fndOrg As String, fndPt As String
Dim Dates_Snap As New ADODB.Recordset
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Private Function chkDuplEEID()
Dim SQLQ As String
Dim snapDupl As New ADODB.Recordset


chkDuplEEID = False

SQLQ = "SELECT * FROM HREMPEQU WHERE ((HREMPEQU.EQ_EMPNBR = " & getEmpnbr(elpEEID.Text)
SQLQ = SQLQ & ") AND (HREMPEQU.EQ_PLAN = '" & clpPlanNbr.Text & "'))"

snapDupl.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapDupl.EOF And snapDupl.BOF Then
    chkDuplEEID = True
Else
    chkDuplEEID = False
End If

End Function

Private Function chkSurveyData()
Dim SQLQ As String
Dim dynNOC As New ADODB.Recordset

chkSurveyData = False

If cmbType.ListIndex = 2 Then
    MsgBox "Type is required!"
    cmbType.SetFocus
    Exit Function
End If

If Len(clpPlanNbr.Text) <= 0 Then
    MsgBox "Plan Number field is required!"
    clpPlanNbr.SetFocus
    Exit Function
Else
  If clpPlanNbr.Caption = "Unassigned" Then
      MsgBox "Plan Number must be valid"
       clpPlanNbr.SetFocus
      Exit Function
  End If
End If

If Len(elpEEID.Text) <= 0 Then
    MsgBox "Employee Number is a required field!"
    elpEEID.SetFocus
    Exit Function
Else
    If elpEEID.Caption = "Unassigned" Then
        MsgBox "Employee Number must be valid"
        elpEEID.SetFocus
        Exit Function
    Else
        If fglbNewRec% = True Then
            If Not chkDuplEEID() Then
                  MsgBox "You have already this Employee in Survey Data. If you want to change this Employee's data, select from the look-up and click Edit!"
                  elpEEID.SetFocus
                  Exit Function
            End If
        End If
    End If
End If

If txtEmpEquity = "" Then
  txtEmpEquity = elpEEID.Text
Else
    If txtEmpEquity.Enabled Then
        If Not IsNumeric(txtEmpEquity) Then
            MsgBox "Employee Equity Number must be numeric"
            txtEmpEquity.SetFocus
            Exit Function
        End If
    End If
End If

If clpPT.Text = "" Then
    clpPT.Text = fndPt
Else
    If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
        MsgBox lStr("Category code must be valid")
         clpPT.SetFocus
        Exit Function
    End If
End If

If Len(dlpSurvDate.Text) <= 0 Then
    Call CR_Dates
    dlpSurvDate.Text = glbSurvDate
Else
    If Not IsDate(dlpSurvDate.Text) Then
        MsgBox "Enter a valid Survey Date!"
        dlpSurvDate.Text = ""
        dlpSurvDate.SetFocus
        Exit Function
    End If
End If
  
If txtSurveyCompleted <> "" Then
    'Ticket #19537
    If glbCompSerial = "S/N - 2279W" Then
        If txtSurveyCompleted <> "Y" And txtSurveyCompleted <> "N" And txtSurveyCompleted <> "R" Then
            MsgBox "Survey Completed must be Y or N or R!"
            txtSurveyCompleted = ""
            txtSurveyCompleted.SetFocus
            Exit Function
        End If
    Else
        If txtSurveyCompleted <> "Y" And txtSurveyCompleted <> "N" Then
            MsgBox "Survey Completed must be Y or N!"
            txtSurveyCompleted = ""
            txtSurveyCompleted.SetFocus
            Exit Function
        End If
    End If
Else
    MsgBox "Survey Completed is a required field !"
    txtSurveyCompleted.SetFocus
    Exit Function
End If

If txtGender <> "" Then
    If txtGender <> "M" And txtGender <> "F" Then
        MsgBox "Gender must be M or F!"
        txtGender = ""
        txtGender.SetFocus
        Exit Function
    End If
End If

If txtAborig <> "" Then
    If txtAborig <> "Y" And txtAborig <> "N" Then
        MsgBox "For Aboriginal you must enter Y or N!!"
        txtAborig = ""
        txtAborig.SetFocus
        Exit Function
    End If
End If

If txtVisibMinority = "" Then
ElseIf txtVisibMinority <> "Y" And txtVisibMinority <> "N" Then
    MsgBox "For Visible Minority you must enter Y or N!"
    txtVisibMinority = ""
    txtVisibMinority.SetFocus
    Exit Function
Else
    If txtAborig = "Y" And txtVisibMinority = "Y" Then
        MsgBox "You have indicated both Aboriginal and Visible Minority. Choose one of the Aboriginal or Visible Minority!"
        txtVisibMinority = ""
        txtVisibMinority.SetFocus
        Exit Function
    End If
End If

If txtDisability <> "" Then
    If txtDisability <> "Y" And txtDisability <> "N" Then
        MsgBox "Disablity must be Y or N!"
        txtDisability = ""
        txtDisability.SetFocus
        Exit Function
    End If
End If

If clpCode(1).Text <> "" Then
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox lStr("Region / CMA code must be valid")
        clpCode(1).Text = ""
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If clpNOGC.Text <> "" Then
    If clpNOGC.Caption = "Unassigned" Then
        MsgBox "N.O.C. code must be valid"
        clpNOGC.Text = ""
        'clpNOGC.SetFocus   'Ticket #24309 - Giving an error as the control is not visible so cannot setfocus
        Exit Function
    End If
Else
    If glbOracle Then
        SQLQ = "SELECT HRJOB.JB_FEDGRP, HR_JOB_HISTORY.JH_EMPNBR, HR_JOB_HISTORY.JH_CURRENT "
        SQLQ = SQLQ & "FROM HRJOB, HR_JOB_HISTORY "
        SQLQ = SQLQ & "WHERE  HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & "AND  HR_JOB_HISTORY.JH_EMPNBR= " & getEmpnbr(elpEEID.Text)
        SQLQ = SQLQ & " AND (HR_JOB_HISTORY.JH_CURRENT<>0)"
    Else
        'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
        If OETYPE.Text = "A" Then
            SQLQ = "SELECT HRJOB.JB_FEDGRP, HR_JOB_HISTORY.JH_EMPNBR, HR_JOB_HISTORY.JH_CURRENT "
            SQLQ = SQLQ & "FROM HRJOB INNER JOIN HR_JOB_HISTORY "
            SQLQ = SQLQ & "ON HRJOB.JB_CODE = HR_JOB_HISTORY.JH_JOB "
            SQLQ = SQLQ & "WHERE ((HR_JOB_HISTORY.JH_EMPNBR= " & getEmpnbr(elpEEID.Text)
            SQLQ = SQLQ & ") AND (HR_JOB_HISTORY.JH_CURRENT<>0))"
        Else
            SQLQ = "SELECT HRJOB.JB_FEDGRP, TERM_JOB_HISTORY.JH_EMPNBR, TERM_JOB_HISTORY.JH_CURRENT "
            SQLQ = SQLQ & "FROM HRJOB INNER JOIN TERM_JOB_HISTORY "
            SQLQ = SQLQ & "ON HRJOB.JB_CODE = TERM_JOB_HISTORY.JH_JOB "
            SQLQ = SQLQ & "WHERE ((TERM_JOB_HISTORY.JH_EMPNBR= " & getEmpnbr(elpEEID.Text)
            SQLQ = SQLQ & ") AND (TERM_JOB_HISTORY.JH_CURRENT<>0))"
        End If
    End If
    dynNOC.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If dynNOC.RecordCount > 0 Then
        If IsNull(dynNOC("JB_FEDGRP")) Then
            clpNOGC.Text = ""
        Else
            clpNOGC.Text = dynNOC("JB_FEDGRP")
        End If
    End If
End If

If Len(dlpDueDate.Text) <= 0 Then
    Call CR_Dates
    dlpDueDate.Text = glbDueDate
Else
    If Not IsDate(dlpDueDate.Text) Then
        MsgBox "Enter a valid Due Date!"
        dlpDueDate.Text = ""
        dlpDueDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpTermDate.Text) > 0 Then
    If Not IsDate(dlpTermDate.Text) Then
        MsgBox "Enter a valid Termination Date!"
        dlpTermDate.Text = ""
        dlpTermDate.SetFocus
        Exit Function
    End If
    If Len(dlpDueDate.Text) > 0 And IsDate(dlpDueDate.Text) Then
        If DaysBetween(dlpDueDate, dlpTermDate.Text) < 0 Then
            MsgBox "Termination date can not be prior to Due Date"
            dlpTermDate.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #25367 - VitalAire
If clpProv.Caption = "Unassigned" Then
    MsgBox "Invalid Province"
    clpProv.SetFocus
    Exit Function
End If


If glbCompSerial = "S/N - 2380W" Then
    If Len(clpCode(6).Text) = 0 Then
        MsgBox lblTitle(18).Caption & " is required field"
        clpCode(6).SetFocus
        Exit Function
    ElseIf Len(clpCode(6).Text) > 0 And clpCode(6).Caption = "Unassigned" Then
        MsgBox lblTitle(18).Caption & " is invalid"
        clpCode(6).SetFocus
        Exit Function
    End If
End If

txtOrg = fndOrg
chkSurveyData = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fglbNew = False

'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
''' Sam add July 2002 * Remove Binding Control
RSDATA.CancelUpdate

Call Display_Value

'Call ST_UPD_MODE(False) ' reset screen's attributes

fndFName = ""
fndLName = ""
fndPt = ""
fndOrg = ""

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HREMPEQU", "Cancel")
Call RollBack '08June99 js

End Sub
'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"

    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'data1.Recordset.Delete
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
''' Sam add July 2002 * Remove Binding Control
gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Me.vbxTrueGrid.SetFocus

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREMPEQU", "Delete")
Call RollBack   '08June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub


'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

fglbNewRec% = True

Call Set_Control("B", Me)
cmdRecal.Enabled = False

RSDATA.AddNew

lblCNum.Caption = "001"
fglbNew = True
Call SET_UP_MODE
'Call ST_UPD_MODE(True)  'May99 js

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRMATRIX", "Add")
Call RollBack '08June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X%

On Error GoTo cmdOK_Err

If Not chkSurveyData() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call elpEEID_Change

Call Set_Control("U", Me, RSDATA)

gdbAdoIhr001.BeginTrans
RSDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

fglbNew = False

Call SET_UP_MODE

'Call ST_UPD_MODE(False) 'May99 js
If fglbEmptyNew Then fglbEmptyNew = False

fglbNewRec% = False
fndFName = ""
fndLName = ""
fndPt = ""
fndOrg = ""

'Hemu 07/02/2003 Begin - Ticket #4247
txtNOC.Text = ""
lblNOCDesc.Caption = ""
'Hemu 07/02/2003 End
clpNOGC.Text = ""   'Ticket #24309

'Ticket #25367 - VitalAire
txtProv.Text = ""

cmdRecal.Enabled = True

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus

Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRMATRIX", "Update")
Call RollBack  '08June99 js
Resume Next
End Sub
'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Survey Data"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

RHeading = "Survey Data"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub


'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub CR_Dates()
Dim SQLQ As String
Dim countr   As Integer

On Error GoTo CR_Dates_Err

Screen.MousePointer = HOURGLASS

glbDueDate = ""
glbSurvDate = ""

SQLQ = "SELECT PP_DUEDATE,PP_SURVEYD FROM HRPARCOP WHERE PP_PLAN = '" & clpPlanNbr.Text & "'"

If Dates_Snap.State <> 0 Then Dates_Snap.Close
Dates_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Not Dates_Snap.EOF Then
  If IsDate(Dates_Snap("PP_DUEDATE")) Then glbDueDate = Dates_Snap("PP_DUEDATE")
  If IsDate(Dates_Snap("PP_SURVEYD")) Then glbSurvDate = Dates_Snap("PP_SURVEYD")
End If

Screen.MousePointer = DEFAULT

Exit Sub

CR_Dates_Err:
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_Dates", "HRPARCOP", "Select")

If gintRollBack% = False Then
    Resume Next
End If

End Sub

Private Sub CR_FLNames()
Dim SQLQ As String
Dim countr   As Integer
On Error GoTo CR_FLNames_Err

Screen.MousePointer = HOURGLASS

If FLNames_Snap.State <> 0 Then FLNames_Snap.Close

SQLQ = "Select ED_PT,ED_ORG,ED_SURNAME,ED_FNAME,ED_REGION,ED_PROV,ED_ORGT1,ED_DIV "


If dlpTermDate.Text = "" And OETYPE = "A" Then
    SQLQ = SQLQ & " FROM HREMP "
    SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
    SQLQ = SQLQ & " AND ED_EMPNBR = " & getEmpnbr(elpEEID.Text)
    FLNames_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic
Else
    SQLQ = SQLQ & " FROM TERM_HREMP "
    SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
    SQLQ = SQLQ & " AND ED_EMPNBR = " & getEmpnbr(elpEEID.Text)
    FLNames_Snap.Open SQLQ, gdbAdoIhr001X, adOpenStatic
End If


fndFName = ""
fndLName = ""
fndPt = ""
fndOrg = ""

If Not FLNames_Snap.EOF Then
    If Not IsNull(FLNames_Snap("ED_FNAME")) Then fndFName = FLNames_Snap("ED_FNAME")
    If Not IsNull(FLNames_Snap("ED_SURNAME")) Then fndLName = FLNames_Snap("ED_SURNAME")
    If Not IsNull(FLNames_Snap("ED_PT")) Then fndPt = FLNames_Snap("ED_PT")
    If Not IsNull(FLNames_Snap("ED_ORG")) Then fndOrg = FLNames_Snap("ED_ORG")
    
    'Ticket #25367 - VitalAire
    If glbCompSerial = "S/N - 2380W" Then
        'Category
        clpPT.Text = fndPt
        
        'Region
        If Not IsNull(FLNames_Snap("ED_REGION")) Then
            clpCode(1).Text = FLNames_Snap("ED_REGION")
        End If
        
        'Province #
        If Not IsNull(FLNames_Snap("ED_PROV")) Then
            clpProv.Text = FLNames_Snap("ED_PROV")
        End If
        
        'CMA - Organizational Code 1
        If Not IsNull(FLNames_Snap("ED_ORGT1")) Then
            clpCode(6).Text = FLNames_Snap("ED_ORGT1")
        End If
        
        'NAICS based on Division
        If Not IsNull(FLNames_Snap("ED_DIV")) Then
            If FLNames_Snap("ED_DIV") = "PWN" Or FLNames_Snap("ED_DIV") = "D2X" Or FLNames_Snap("ED_DIV") = "MZT" Then
                txtNAICS.Text = "62161"
            ElseIf FLNames_Snap("ED_DIV") = "XTM" Or FLNames_Snap("ED_DIV") = "XTT" Then
                txtNAICS.Text = "236220"
            End If
        End If
    Else
        'Province #
        If Not IsNull(FLNames_Snap("ED_PROV")) Then
            clpProv.Text = FLNames_Snap("ED_PROV")
        End If
    
        'CMA - Organizational Code 1
        If Not IsNull(FLNames_Snap("ED_ORGT1")) Then
            clpCode(6).Text = FLNames_Snap("ED_ORGT1")
        End If
    End If
    
    'Hemu 07/02/2003 Begin - Ticket #4247 Display the NOC Code and Description
    Dim rsJobHis As New ADODB.Recordset
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    If OETYPE = "A" Then
        rsJobHis.Open "SELECT OC_CODE, OC_SDESCR FROM HR_OCCUPATION_CLASS WHERE OC_CODE = (SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT TOP 1 JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & getEmpnbr(elpEEID.Text) & " AND JH_CURRENT <> 0))", gdbAdoIhr001, adOpenStatic
    Else
        rsJobHis.Open "SELECT OC_CODE, OC_SDESCR FROM HR_OCCUPATION_CLASS WHERE OC_CODE = (SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT TOP 1 JH_JOB FROM TERM_JOB_HISTORY WHERE JH_EMPNBR = " & getEmpnbr(elpEEID.Text) & " AND JH_CURRENT <> 0 ORDER BY TERM_SEQ DESC))", gdbAdoIhr001, adOpenStatic
    End If
    If Not rsJobHis.EOF Then
        txtNOC.Text = rsJobHis("OC_CODE")
        clpNOGC.Text = rsJobHis("OC_CODE")  'Ticket #24309
        lblNOCDesc.Caption = rsJobHis("OC_SDESCR")
    Else
        txtNOC.Text = ""
        clpNOGC.Text = ""       'Ticket #24309
        lblNOCDesc.Caption = ""
    End If
    rsJobHis.Close
    'Hemu 07/02/2003 End
End If
Screen.MousePointer = DEFAULT
Exit Sub

CR_FLNames_Err:
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_FLNames", "HREMP", "Select")

If gintRollBack% = False Then
    Resume Next
End If

End Sub

'Hemu 07/02/2003  Begin - Ticket #4247
Private Sub Display_NOC_Desc(NOC_Code)
    Dim rsJobHis As New ADODB.Recordset
    rsJobHis.Open "SELECT OC_SDESCR FROM HR_OCCUPATION_CLASS WHERE OC_CODE = '" & NOC_Code & "'", gdbAdoIhr001, adOpenStatic
    If Not rsJobHis.EOF Then
        lblNOCDesc.Caption = rsJobHis("OC_SDESCR")
    Else
        lblNOCDesc.Caption = ""
    End If
    rsJobHis.Close
End Sub
'Hemu 07/02/2003  Begin - Ticket #4247

Private Sub clpPlanNbr_Change()
    'Retrieve Questions of the Plan #
    Call Retrieve_Plan_Questions
End Sub

Private Sub clpPlanNbr_LostFocus()
    'Retrieve Questions of the Plan #
    'Call Retrieve_Plan_Questions
End Sub

Private Sub clpProv_Change()
    If glbCompSerial = "S/N - 2380W" Then   'VitalAire
        txtProv.Text = Get_ProvinceCodeData(clpProv, "NBR")
    End If
End Sub

Private Sub cmbType_Change()
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    Call cmbType_Click
End Sub

Private Sub cmbType_Click()
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    If cmbType.ListIndex = 0 Then
        'elpEEID.ShowDescription = True
        OETYPE = "A"
        elpEEID.LookupType = 0
    ElseIf cmbType.ListIndex = 1 Then
        OETYPE = "T"
        elpEEID.LookupType = 1
    End If
End Sub

Private Sub cmbType_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Hemu 07/02/2003  Begin - Ticket #4247, Re-updating all the Employment Equity data with NOC Code
'                        and Termination Date
Private Sub cmdRecal_Click()
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 1
    MDIMain.panHelp(0).FloodPercent = 3

    Call UpdateHREMPEQU_NOGC
    Call InputHREMPEQU_DOT
    Call UpdHREMPEQU_Type
    
    Data1.Refresh
    
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
End Sub
'Hemu 07/02/2003  Begin - Ticket #4247

'Hemu 07/02/2003  Begin - Ticket #4247, Update NOC Code
Private Sub UpdateHREMPEQU_NOGC()
On Error GoTo UpdateHREMPEQU_NOGC_Err
Dim rsEmpEQU As New ADODB.Recordset
Dim rsEmpNOC As New ADODB.Recordset
Dim SQLQ As String
Dim dblPerc, FloodPerc As Double
    
    SQLQ = "SELECT * FROM HREMPEQU"
    rsEmpEQU.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsEmpEQU.EOF Then
        rsEmpEQU.MoveFirst
        
        dblPerc = (50 / rsEmpEQU.RecordCount)
        FloodPerc = dblPerc
        
        gdbAdoIhr001.BeginTrans
        Do While Not rsEmpEQU.EOF
            'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
            If rsEmpEQU("EQ_TYPE") = "A" Then
                rsEmpNOC.Open "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT TOP 1 JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & rsEmpEQU("EQ_EMPNBR") & " AND JH_CURRENT <> 0)", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Else
                rsEmpNOC.Open "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT TOP 1 JH_JOB FROM TERM_JOB_HISTORY WHERE JH_EMPNBR = " & rsEmpEQU("EQ_EMPNBR") & " AND JH_CURRENT <> 0)", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            End If
            
            If Not rsEmpNOC.EOF Then
                If Not IsNull(rsEmpNOC("JB_FEDGRP")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NOGC = '" & rsEmpNOC("JB_FEDGRP") & "' WHERE EQ_EMPNBR = " & rsEmpEQU("EQ_EMPNBR")
                End If
            End If
            rsEmpEQU.MoveNext
            
            MDIMain.panHelp(0).FloodPercent = FloodPerc
            FloodPerc = FloodPerc + dblPerc
            
            rsEmpNOC.Close
        Loop
        gdbAdoIhr001.CommitTrans
        
    End If
    rsEmpEQU.Close
    MDIMain.panHelp(0).FloodPercent = 50
    Exit Sub
    
    
UpdateHREMPEQU_NOGC_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateHREMPEQU_NOGC", "HREMPEQU", "Update")
Call RollBack  '08June99 js

End Sub
'Hemu 07/02/2003  Begin - Ticket #4247

'Hemu 07/02/2003  Begin - Ticket #4247, Update Date of Termination
Private Sub InputHREMPEQU_DOT()
On Error GoTo InputHREMPEQU_DOT_Err

Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset
Dim rsTermEmp As New ADODB.Recordset
Dim dblPerc, FloodPerc As Double
Dim xDiv As String
Dim xProv As String
Dim xRegion As String

SQLQ = "SELECT * FROM HREMPEQU"
dynEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Not dynEmp.EOF Then
    dynEmp.MoveFirst
    
    dblPerc = (50 / dynEmp.RecordCount)
    FloodPerc = 50 + dblPerc
    
    gdbAdoIhr001.BeginTrans
    Do While Not dynEmp.EOF
        rsTermEmp.Open "SELECT Term_DOT, TERM_SEQ FROM TERM_HRTRMEMP WHERE Employee_Number = " & dynEmp("EQ_EMPNBR"), gdbAdoIhr001X, adOpenStatic
        
        If Not rsTermEmp.EOF Then
            'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
            'gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_DOT = " & Date_SQL(rsTermEmp("Term_DOT")) & " WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
            gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_DOT = " & Date_SQL(rsTermEmp("Term_DOT")) & ", EQ_TYPE = 'T' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
            
            'Ticket #25367 - VitalAire
            If glbCompSerial = "S/N - 2380W" Then
                'Get Terminated employee's division
                xDiv = GetTermEmpData(dynEmp("EQ_EMPNBR"), rsTermEmp("TERM_SEQ"), "ED_DIV", "")
                If (xDiv = "PWN" Or xDiv = "D2X" Or xDiv = "MZT") And IsNull(dynEmp("EQ_NAICS")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NAICS = '62161' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                ElseIf (xDiv = "XTM" Or xDiv = "XTT") And IsNull(dynEmp("EQ_NAICS")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NAICS = '236220' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                End If
                
                If IsNull(dynEmp("EQ_PROV")) Then
                    xProv = GetTermEmpData(dynEmp("EQ_EMPNBR"), rsTermEmp("TERM_SEQ"), "ED_PROV", "")
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_PROV = '" & Get_ProvinceCodeData(xProv, "NBR") & "' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                End If
            
                If IsNull(dynEmp("EQ_REGION")) Then
                    xRegion = GetTermEmpData(dynEmp("EQ_EMPNBR"), rsTermEmp("TERM_SEQ"), "ED_REGION", "")
                    If xRegion <> "" Then
                        gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_REGION = '" & xRegion & "' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                    End If
                End If
            End If
        End If
        dynEmp.MoveNext
        
        MDIMain.panHelp(0).FloodPercent = FloodPerc
        FloodPerc = FloodPerc + dblPerc
        
        rsTermEmp.Close
    Loop
    gdbAdoIhr001.CommitTrans
    
End If
dynEmp.Close
Exit Sub

InputHREMPEQU_DOT_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InputHREMPEQU_DOT", "TERM_HRTRMEMP", "Update")
Call RollBack  '08June99 js

End Sub
'Hemu 07/02/2003 Begin - Ticket #4247

Private Sub UpdHREMPEQU_Type()
On Error GoTo UpdHREMPEQU_Type_Err

Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim dblPerc, FloodPerc As Double

SQLQ = "SELECT * FROM HREMPEQU"
dynEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Not dynEmp.EOF Then
    dynEmp.MoveFirst
    
    dblPerc = (50 / dynEmp.RecordCount)
    FloodPerc = 50 + dblPerc
    
    gdbAdoIhr001.BeginTrans
    Do While Not dynEmp.EOF
        rsHREmp.Open "SELECT ED_EMPNBR, ED_DIV, ED_PROV, ED_REGION FROM HREMP WHERE ED_EMPNBR = " & dynEmp("EQ_EMPNBR"), gdbAdoIhr001, adOpenStatic
        
        If Not rsHREmp.EOF Then
            gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_TYPE = 'A' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
            
            'Ticket #25367 - VitalAire
            If glbCompSerial = "S/N - 2380W" Then
                If (rsHREmp("ED_DIV") = "PWN" Or rsHREmp("ED_DIV") = "D2X" Or rsHREmp("ED_DIV") = "MZT") And IsNull(dynEmp("EQ_NAICS")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NAICS = '62161' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                ElseIf (rsHREmp("ED_DIV") = "XTM" Or rsHREmp("ED_DIV") = "XTT") And IsNull(dynEmp("EQ_NAICS")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_NAICS = '236220' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                End If
                
                If IsNull(dynEmp("EQ_PROV")) And Not IsNull(rsHREmp("ED_PROV")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_PROV = '" & Get_ProvinceCodeData(rsHREmp("ED_PROV"), "NBR") & "' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                End If
                
                If IsNull(dynEmp("EQ_REGION")) And Not IsNull(rsHREmp("ED_REGION")) Then
                    gdbAdoIhr001.Execute "UPDATE HREMPEQU SET EQ_REGION = '" & rsHREmp("ED_REGION") & "' WHERE EQ_EMPNBR = " & dynEmp("EQ_EMPNBR")
                End If
            End If
        End If
        dynEmp.MoveNext
        
        MDIMain.panHelp(0).FloodPercent = FloodPerc
        FloodPerc = FloodPerc + dblPerc
        
        rsHREmp.Close
    Loop
    gdbAdoIhr001.CommitTrans
    
End If
dynEmp.Close
Exit Sub

UpdHREMPEQU_Type_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdHREMPEQU_Type", "HRMEMP", "Update")
Call RollBack  '08June99 js

End Sub
'Hemu 07/02/2003 Begin - Ticket #4247

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "PAYROLL", "SELECT")

End Sub


Private Sub elpEEID_Change()

If Len(elpEEID) > 0 Then
    Call cmbType_Click
    Call CR_FLNames
    txtFirstName = fndFName
    txtSurname = fndLName
End If
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%, X%

glbOnTop = Me.name

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
cmbType.Clear
cmbType.AddItem "Active"
cmbType.AddItem "Terminated"
cmbType.AddItem ""


If glbCompSerial = "S/N - 2227W" Then
     clpCode(1).MaxLength = 6
End If

'Ticket #19537
If glbCompSerial = "S/N - 2279W" Then
    txtSurveyCompleted.Tag = "01-Enter Y / N / R"
End If

'Ticket #25367 - VitalAire
If glbCompSerial = "S/N - 2380W" Then
    lblTitle(18).FontBold = True
    lblRegion.Caption = "Region"
    
    lblTitle(1).Visible = True
    txtNAICS.Visible = True
    
    clpProv.DataField = ""
    txtProv.DataField = "EQ_PROV"
    txtProv.Visible = True
    clpProv.Left = 2485
Else
    lblTitle(1).Visible = False
    txtNAICS.Visible = False
    
    clpProv.DataField = "EQ_PROV"
    txtProv.DataField = ""
    
    txtProv.Visible = False
    clpProv.Left = clpPT.Left
End If

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HREMPEQU ORDER BY EQ_SURNAME"
Data1.Refresh
Me.Show

Screen.MousePointer = DEFAULT

Call Display_Value
'Call ST_UPD_MODE(False)                                   '

If Not gSec_Matrix Then                                    'May99 js
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False                             '
End If

Call setCaption(lblPT)
Call setCaption(lblRegion)
Call setCaption(lblTitle(18))

clpPlanNbr.TextBoxWidth = 1200

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT                           '

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

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmSurveyData = Nothing
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF              'May99 js
'cmdCancel.Enabled = TF          '
'cmdClose.Enabled = FT           '
'cmdModify.Enabled = FT          '
'cmdNew.Enabled = FT             '
'cmdDelete.Enabled = FT          '
'cmdPrint.Enabled = FT           '

'Hemu 07/02/2003 Begin - Ticket #4247
'cmdRecal.Enabled = FT
'Hemu 07/02/2003 End

txtAborig.Enabled = TF          '
clpCode(1).Enabled = TF         '
txtDisability.Enabled = TF      '
dlpDueDate.Enabled = TF         '
elpEEID.Enabled = TF            '
txtEmpEquity.Enabled = TF       '
txtGender.Enabled = TF          '
clpNOGC.Enabled = TF            '
clpPlanNbr.Enabled = TF         '
clpPT.Enabled = TF              '
dlpSurvDate.Enabled = TF        '
txtSurveyCompleted.Enabled = TF '
dlpTermDate.Enabled = TF        '
txtVisibMinority.Enabled = TF   '
If Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
cmbType.Enabled = TF

'Ticket #25367 - VitalAire, For all
'If glbCompSerial = "S/N - 2380W" Then
    clpProv.Enabled = TF
    clpCode(6).Enabled = TF
    
    If glbCompSerial = "S/N - 2380W" Then
        txtNAICS.Enabled = TF
    End If
'End If

End Sub

Private Sub OETYPE_Change()
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    If Len(OETYPE) > 0 Then
        If OETYPE.Text = "A" Then
            cmbType.ListIndex = 0
            elpEEID.LookupType = 0
        ElseIf OETYPE.Text = "T" Then
            cmbType.ListIndex = 1
            elpEEID.LookupType = 1
        Else
            cmbType.ListIndex = 2   '-1
            elpEEID.LookupType = 0
        End If
    Else
        cmbType.ListIndex = 2   '-1
        elpEEID.LookupType = 0
    End If
        
End Sub

Private Sub optQ1YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ1YesNo_KeyPress(Index As Integer, KeyAscii As Integer)
    'Set Answer to a text field
    If optQ1YesNo(0) = True Then
        txtQ1.Text = "Y"
    ElseIf optQ1YesNo(1) = True Then
        txtQ1.Text = "N"
    End If
End Sub

Private Sub optQ1YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ1YesNo(0) = True Then
        txtQ1.Text = "Y"
    ElseIf optQ1YesNo(1) = True Then
        txtQ1.Text = "N"
    End If
End Sub

Private Sub optQ2YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ2YesNo(0) = True Then
        txtQ2.Text = "Y"
    ElseIf optQ2YesNo(1) = True Then
        txtQ2.Text = "N"
    End If
End Sub

Private Sub optQ2YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ2YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ2YesNo(0) = True Then
        txtQ2.Text = "Y"
    ElseIf optQ2YesNo(1) = True Then
        txtQ2.Text = "N"
    End If
End Sub

Private Sub optQ3YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ3YesNo(0).Value = True Then
        txtQ3.Text = "Y"
    ElseIf optQ3YesNo(1).Value = True Then
        txtQ3.Text = "N"
    End If
End Sub

Private Sub optQ3YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ3YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ3YesNo(0).Value = True Then
        txtQ3.Text = "Y"
    ElseIf optQ3YesNo(1).Value = True Then
        txtQ3.Text = "N"
    End If
End Sub

Private Sub optQ4YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ4YesNo(0).Value = True Then
        txtQ4.Text = "Y"
    ElseIf optQ4YesNo(1).Value = True Then
        txtQ4.Text = "N"
    End If
End Sub

Private Sub optQ4YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ4YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ4YesNo(0).Value = True Then
        txtQ4.Text = "Y"
    ElseIf optQ4YesNo(1).Value = True Then
        txtQ4.Text = "N"
    End If
End Sub

Private Sub optQ5YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ5YesNo(0).Value = True Then
        txtQ5.Text = "Y"
    ElseIf optQ5YesNo(1).Value = True Then
        txtQ5.Text = "N"
    End If
End Sub

Private Sub optQ5YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ5YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ5YesNo(0).Value = True Then
        txtQ5.Text = "Y"
    ElseIf optQ5YesNo(1).Value = True Then
        txtQ5.Text = "N"
    End If
End Sub

Private Sub optQ6YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ6YesNo(0).Value = True Then
        txtQ6.Text = "Y"
    ElseIf optQ6YesNo(1).Value = True Then
        txtQ6.Text = "N"
    End If
End Sub

Private Sub optQ6YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ6YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ6YesNo(0).Value = True Then
        txtQ6.Text = "Y"
    ElseIf optQ6YesNo(1).Value = True Then
        txtQ6.Text = "N"
    End If
End Sub

Private Sub optQ7YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ7YesNo(0).Value = True Then
        txtQ7.Text = "Y"
    ElseIf optQ7YesNo(1).Value = True Then
        txtQ7.Text = "N"
    End If
End Sub

Private Sub optQ7YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ7YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ7YesNo(0).Value = True Then
        txtQ7.Text = "Y"
    ElseIf optQ7YesNo(1).Value = True Then
        txtQ7.Text = "N"
    End If
End Sub

Private Sub optQ8YesNo_KeyPress(Index As Integer, Value As Integer)
    'Set Answer to a text field
    If optQ8YesNo(0).Value = True Then
        txtQ8.Text = "Y"
    ElseIf optQ8YesNo(1).Value = True Then
        txtQ8.Text = "N"
    End If
End Sub

Private Sub optQ8YesNo_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optQ8YesNo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Set Answer to a text field
    If optQ8YesNo(0).Value = True Then
        txtQ8.Text = "Y"
    ElseIf optQ8YesNo(1).Value = True Then
        txtQ8.Text = "N"
    End If
End Sub

Private Sub txtAborig_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtAborig_LostFocus()
    'Hemu - 07/02/2003 Begin - Jerry suggested
    txtAborig.Text = UCase(txtAborig.Text)
    'Hemu - 07/02/2003 End
End Sub

Private Sub txtDisability_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDisability_LostFocus()
    'Hemu - 07/02/2003 Begin - Jerry suggested
    txtDisability.Text = UCase(txtDisability.Text)
    'Hemu - 07/02/2003 End
End Sub

Private Sub txtEmpEquity_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtGender_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtGender_LostFocus()
    'Hemu - 07/02/2003 Begin - Jerry suggested
    txtGender.Text = UCase(txtGender.Text)
    'Hemu - 07/02/2003 End
End Sub

Private Sub txtProv_Change()
    If glbCompSerial = "S/N - 2380W" Then   'VitalAire
        clpProv.Text = Get_ProvinceNoData(txtProv.Text, "CODE")
    End If
End Sub

Private Sub txtQ1_Change()
    If txtQ1.Text = "Y" Then
        optQ1YesNo(0) = True
        optQ1YesNo(1) = False
    Else
        optQ1YesNo(1) = True
        optQ1YesNo(0) = False
    End If
End Sub

Private Sub txtQ2_Change()
    If txtQ2.Text = "Y" Then
        optQ2YesNo(0) = True
        optQ2YesNo(1) = False
    Else
        optQ2YesNo(1) = True
        optQ2YesNo(0) = False
    End If
End Sub

Private Sub txtQ3_Change()
    If txtQ3.Text = "Y" Then
        optQ3YesNo(0) = True
        optQ3YesNo(1) = False
    Else
        optQ3YesNo(1) = True
        optQ3YesNo(0) = False
    End If
End Sub

Private Sub txtQ4_Change()
    If txtQ4.Text = "Y" Then
        optQ4YesNo(0) = True
        optQ4YesNo(1) = False
    Else
        optQ4YesNo(1) = True
        optQ4YesNo(0) = False
    End If
End Sub

Private Sub txtQ5_Change()
    If txtQ5.Text = "Y" Then
        optQ5YesNo(0) = True
        optQ5YesNo(1) = False
    Else
        optQ5YesNo(1) = True
        optQ5YesNo(0) = False
    End If
End Sub

Private Sub txtQ6_Change()
    If txtQ6.Text = "Y" Then
        optQ6YesNo(0) = True
        optQ6YesNo(1) = False
    Else
        optQ6YesNo(1) = True
        optQ6YesNo(0) = False
    End If
End Sub

Private Sub txtQ7_Change()
    If txtQ7.Text = "Y" Then
        optQ7YesNo(0) = True
        optQ7YesNo(1) = False
    Else
        optQ7YesNo(1) = True
        optQ7YesNo(0) = False
    End If
End Sub

Private Sub txtQ8_Change()
    If txtQ8.Text = "Y" Then
        optQ8YesNo(0) = True
        optQ8YesNo(1) = False
    Else
        optQ8YesNo(1) = True
        optQ8YesNo(0) = False
    End If
End Sub

Private Sub txtSurveyCompleted_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSurveyCompleted_LostFocus()
    'Hemu - 07/02/2003 Begin - Jerry suggested
    txtSurveyCompleted.Text = UCase(txtSurveyCompleted.Text)
    'Hemu - 07/02/2003 End
End Sub

Private Sub txtVisibMinority_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtVisibMinority_LostFocus()
    'Hemu - 07/02/2003 Begin - Jerry suggested
    txtVisibMinority.Text = UCase(txtVisibMinority.Text)
    'Hemu - 07/02/2003 End
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
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


''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        If glbtermopen Then
            RSDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HREMPEQU where EQ_ID= " & Data1.Recordset!EQ_ID & " ORDER BY EQ_SURNAME"
    
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    
    Call SET_UP_MODE

End Sub


Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
       
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HREMPEQU"
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'Hemu 07/02/2003 Begin - Ticket #4247
lblNOCDesc.Caption = ""
'Hemu 07/02/2003 End

Call Display_Value
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
UpdateRight = gSec_Upd_EmploymentEQT 'True
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
ElseIf RSDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Private Sub txtNOC_Change()
'Hemu 07/02/2003 Begin - Ticket #4247
    If Len(txtNOC.Text) > 0 Then
        Call Display_NOC_Desc(txtNOC.Text)
    End If
'Hemu 07/02/2003 End
End Sub

Private Sub Retrieve_Plan_Questions()
    Dim rsHRParCop As New ADODB.Recordset
    Dim SQLQ As String
    
    If Trim(clpPlanNbr.Text) = "" Then
        txtQues1.Text = ""
        txtQues2.Text = ""
        txtQues3.Text = ""
        txtQues4.Text = ""
        txtQues5.Text = ""
        txtQues6.Text = ""
        txtQues7.Text = ""
        txtQues8.Text = ""
    Else
        SQLQ = "SELECT * FROM HRPARCOP WHERE PP_PLAN = '" & clpPlanNbr.Text & "'"
        rsHRParCop.Open SQLQ, gdbAdoIhr001, adOpenStatic
        
        If Not rsHRParCop.EOF Then
            txtQues1.Text = IIf(IsNull(rsHRParCop("PP_Q1")), "", rsHRParCop("PP_Q1"))
            txtQues2.Text = IIf(IsNull(rsHRParCop("PP_Q2")), "", rsHRParCop("PP_Q2"))
            txtQues3.Text = IIf(IsNull(rsHRParCop("PP_Q3")), "", rsHRParCop("PP_Q3"))
            txtQues4.Text = IIf(IsNull(rsHRParCop("PP_Q4")), "", rsHRParCop("PP_Q4"))
            txtQues5.Text = IIf(IsNull(rsHRParCop("PP_Q5")), "", rsHRParCop("PP_Q5"))
            txtQues6.Text = IIf(IsNull(rsHRParCop("PP_Q6")), "", rsHRParCop("PP_Q6"))
            txtQues7.Text = IIf(IsNull(rsHRParCop("PP_Q7")), "", rsHRParCop("PP_Q7"))
            txtQues8.Text = IIf(IsNull(rsHRParCop("PP_Q8")), "", rsHRParCop("PP_Q8"))
        Else
            txtQues1.Text = ""
            txtQues2.Text = ""
            txtQues3.Text = ""
            txtQues4.Text = ""
            txtQues5.Text = ""
            txtQues6.Text = ""
            txtQues7.Text = ""
            txtQues8.Text = ""
        End If
        rsHRParCop.Close
    End If

End Sub

