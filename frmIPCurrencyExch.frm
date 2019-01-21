VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmIPCurrencyExch 
   Caption         =   "Currency Exchange Table"
   ClientHeight    =   11340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11340
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   19
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   177
      Tag             =   "01-Country"
      Top             =   9720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   19
      Left            =   5880
      TabIndex        =   176
      Tag             =   "00-Country"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   18
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   173
      Tag             =   "01-Country"
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   18
      Left            =   5880
      TabIndex        =   172
      Tag             =   "00-Country"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   17
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   169
      Tag             =   "01-Country"
      Top             =   9000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   17
      Left            =   5880
      TabIndex        =   168
      Tag             =   "00-Country"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   16
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   165
      Tag             =   "01-Country"
      Top             =   8640
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   16
      Left            =   5880
      TabIndex        =   164
      Tag             =   "00-Country"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   15
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   161
      Tag             =   "01-Country"
      Top             =   8280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   15
      Left            =   5880
      TabIndex        =   160
      Tag             =   "00-Country"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   14
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   157
      Tag             =   "01-Country"
      Top             =   7920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   14
      Left            =   5880
      TabIndex        =   156
      Tag             =   "00-Country"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   13
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   153
      Tag             =   "01-Country"
      Top             =   7560
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   13
      Left            =   5880
      TabIndex        =   152
      Tag             =   "00-Country"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   12
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   149
      Tag             =   "01-Country"
      Top             =   7200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   12
      Left            =   5880
      TabIndex        =   148
      Tag             =   "00-Country"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   11
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   145
      Tag             =   "01-Country"
      Top             =   6840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   11
      Left            =   5880
      TabIndex        =   144
      Tag             =   "00-Country"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   10
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   141
      Tag             =   "01-Country"
      Top             =   6480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   5880
      TabIndex        =   140
      Tag             =   "00-Country"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   137
      Tag             =   "01-Country"
      Top             =   6120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   9
      Left            =   5880
      TabIndex        =   136
      Tag             =   "00-Country"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   133
      Tag             =   "01-Country"
      Top             =   5760
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   8
      Left            =   5880
      TabIndex        =   132
      Tag             =   "00-Country"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   129
      Tag             =   "01-Country"
      Top             =   5400
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   5880
      TabIndex        =   128
      Tag             =   "00-Country"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   125
      Tag             =   "01-Country"
      Top             =   5040
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   5880
      TabIndex        =   124
      Tag             =   "00-Country"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   121
      Tag             =   "01-Country"
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   5880
      TabIndex        =   120
      Tag             =   "00-Country"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   117
      Tag             =   "01-Country"
      Top             =   4320
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   5880
      TabIndex        =   116
      Tag             =   "00-Country"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   113
      Tag             =   "01-Country"
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   5880
      TabIndex        =   112
      Tag             =   "00-Country"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   109
      Tag             =   "01-Country"
      Top             =   3600
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   5880
      TabIndex        =   108
      Tag             =   "00-Country"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   105
      Tag             =   "01-Country"
      Top             =   3240
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   5880
      TabIndex        =   104
      Tag             =   "00-Country"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtCountry2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   101
      Tag             =   "01-Country"
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.ComboBox comCountry2 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   5880
      TabIndex        =   100
      Tag             =   "00-Country"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   19
      Left            =   120
      TabIndex        =   61
      Tag             =   "00-Country"
      Top             =   9720
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   18
      Left            =   120
      TabIndex        =   58
      Tag             =   "00-Country"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   17
      Left            =   120
      TabIndex        =   55
      Tag             =   "00-Country"
      Top             =   9000
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   16
      Left            =   120
      TabIndex        =   52
      Tag             =   "00-Country"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   15
      Left            =   120
      TabIndex        =   49
      Tag             =   "00-Country"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   14
      Left            =   120
      TabIndex        =   46
      Tag             =   "00-Country"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   13
      Left            =   120
      TabIndex        =   43
      Tag             =   "00-Country"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   12
      Left            =   120
      TabIndex        =   40
      Tag             =   "00-Country"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   11
      Left            =   120
      TabIndex        =   37
      Tag             =   "00-Country"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   10
      Left            =   120
      TabIndex        =   34
      Tag             =   "00-Country"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   9
      Left            =   120
      TabIndex        =   31
      Tag             =   "00-Country"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   8
      Left            =   120
      TabIndex        =   28
      Tag             =   "00-Country"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Tag             =   "00-Country"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Tag             =   "00-Country"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Tag             =   "00-Country"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Tag             =   "00-Country"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Tag             =   "00-Country"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Tag             =   "00-Country"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Tag             =   "00-Country"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox comCountry1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Tag             =   "00-Country"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ComboBox ComMTH 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "01-Month"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtMTH 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5760
      MaxLength       =   3
      TabIndex        =   72
      Text            =   "MTH"
      Top             =   1800
      Visible         =   0   'False
      Width           =   570
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmIPCurrencyExch.frx":0000
      Height          =   1455
      Left            =   120
      OleObjectBlob   =   "frmIPCurrencyExch.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   525
      Left            =   0
      TabIndex        =   67
      Top             =   10815
      Width           =   13380
      _Version        =   65536
      _ExtentX        =   23601
      _ExtentY        =   926
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
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   600
         TabIndex        =   71
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   4440
         TabIndex        =   70
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CommandButton cmdDeleteEnt 
         Appearance      =   0  'Flat
         Caption         =   "&Delete Entitlement"
         Height          =   375
         Left            =   2520
         TabIndex        =   69
         Tag             =   "Delete all matching records to the above"
         Top             =   120
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   6360
         TabIndex        =   68
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   405
         Left            =   9240
         Top             =   0
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Tag             =   "00-Currency Indicator - Code"
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   7200
      TabIndex        =   64
      Tag             =   "00-Currency Indicator - Code"
      Top             =   2280
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox MskFiscalYear 
      DataField       =   "IP_YEAR"
      Height          =   315
      Left            =   1995
      TabIndex        =   1
      Tag             =   "01-High Dollars"
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###0"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   0
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   74
      Tag             =   "01-Country"
      Top             =   2880
      Visible         =   0   'False
      Width           =   555
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   5
      Tag             =   "00-Currency Indicator - Code"
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   6
      Tag             =   "21-Enter Exchange Rate"
      Top             =   2880
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Tag             =   "00-Currency Indicator - Code"
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Tag             =   "21-Enter Exchange Rate"
      Top             =   3240
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   81
      Tag             =   "01-Country"
      Top             =   3240
      Visible         =   0   'False
      Width           =   555
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   11
      Tag             =   "00-Currency Indicator - Code"
      Top             =   3600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   12
      Tag             =   "21-Enter Exchange Rate"
      Top             =   3600
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   14
      Tag             =   "00-Currency Indicator - Code"
      Top             =   3960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   15
      Tag             =   "21-Enter Exchange Rate"
      Top             =   3960
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   4
      Left            =   1680
      TabIndex        =   17
      Tag             =   "00-Currency Indicator - Code"
      Top             =   4320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   18
      Tag             =   "21-Enter Exchange Rate"
      Top             =   4320
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   20
      Tag             =   "00-Currency Indicator - Code"
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   21
      Tag             =   "21-Enter Exchange Rate"
      Top             =   4680
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   6
      Left            =   1680
      TabIndex        =   23
      Tag             =   "00-Currency Indicator - Code"
      Top             =   5040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   24
      Tag             =   "21-Enter Exchange Rate"
      Top             =   5040
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   7
      Left            =   1680
      TabIndex        =   26
      Tag             =   "00-Currency Indicator - Code"
      Top             =   5400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   7
      Left            =   4680
      TabIndex        =   27
      Tag             =   "21-Enter Exchange Rate"
      Top             =   5400
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   82
      Tag             =   "01-Country"
      Top             =   3600
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   83
      Tag             =   "01-Country"
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   84
      Tag             =   "01-Country"
      Top             =   4320
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   85
      Tag             =   "01-Country"
      Top             =   4680
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   86
      Tag             =   "01-Country"
      Top             =   5040
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   87
      Tag             =   "01-Country"
      Top             =   5400
      Visible         =   0   'False
      Width           =   555
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   8
      Left            =   1680
      TabIndex        =   29
      Tag             =   "00-Currency Indicator - Code"
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   30
      Tag             =   "21-Enter Exchange Rate"
      Top             =   5760
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   9
      Left            =   1680
      TabIndex        =   32
      Tag             =   "00-Currency Indicator - Code"
      Top             =   6120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   33
      Tag             =   "21-Enter Exchange Rate"
      Top             =   6120
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   10
      Left            =   1680
      TabIndex        =   35
      Tag             =   "00-Currency Indicator - Code"
      Top             =   6480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   10
      Left            =   4680
      TabIndex        =   36
      Tag             =   "21-Enter Exchange Rate"
      Top             =   6480
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   11
      Left            =   1680
      TabIndex        =   38
      Tag             =   "00-Currency Indicator - Code"
      Top             =   6840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   11
      Left            =   4680
      TabIndex        =   39
      Tag             =   "21-Enter Exchange Rate"
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   12
      Left            =   1680
      TabIndex        =   41
      Tag             =   "00-Currency Indicator - Code"
      Top             =   7200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   12
      Left            =   4680
      TabIndex        =   42
      Tag             =   "21-Enter Exchange Rate"
      Top             =   7200
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   13
      Left            =   1680
      TabIndex        =   44
      Tag             =   "00-Currency Indicator - Code"
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   13
      Left            =   4680
      TabIndex        =   45
      Tag             =   "21-Enter Exchange Rate"
      Top             =   7560
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   14
      Left            =   1680
      TabIndex        =   47
      Tag             =   "00-Currency Indicator - Code"
      Top             =   7920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   14
      Left            =   4680
      TabIndex        =   48
      Tag             =   "21-Enter Exchange Rate"
      Top             =   7920
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   15
      Left            =   1680
      TabIndex        =   50
      Tag             =   "00-Currency Indicator - Code"
      Top             =   8280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   15
      Left            =   4680
      TabIndex        =   51
      Tag             =   "21-Enter Exchange Rate"
      Top             =   8280
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   16
      Left            =   1680
      TabIndex        =   53
      Tag             =   "00-Currency Indicator - Code"
      Top             =   8640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   16
      Left            =   4680
      TabIndex        =   54
      Tag             =   "21-Enter Exchange Rate"
      Top             =   8640
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   17
      Left            =   1680
      TabIndex        =   56
      Tag             =   "00-Currency Indicator - Code"
      Top             =   9000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   17
      Left            =   4680
      TabIndex        =   57
      Tag             =   "21-Enter Exchange Rate"
      Top             =   9000
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   18
      Left            =   1680
      TabIndex        =   59
      Tag             =   "00-Currency Indicator - Code"
      Top             =   9360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   18
      Left            =   4680
      TabIndex        =   60
      Tag             =   "21-Enter Exchange Rate"
      Top             =   9360
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode1 
      Height          =   285
      Index           =   19
      Left            =   1680
      TabIndex        =   62
      Tag             =   "00-Currency Indicator - Code"
      Top             =   9720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate1 
      Height          =   285
      Index           =   19
      Left            =   4680
      TabIndex        =   63
      Tag             =   "21-Enter Exchange Rate"
      Top             =   9720
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   88
      Tag             =   "01-Country"
      Top             =   5760
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   9
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   89
      Tag             =   "01-Country"
      Top             =   6120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   10
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   90
      Tag             =   "01-Country"
      Top             =   6480
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   11
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   91
      Tag             =   "01-Country"
      Top             =   6840
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   12
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   92
      Tag             =   "01-Country"
      Top             =   7200
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   13
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   93
      Tag             =   "01-Country"
      Top             =   7560
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   14
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   94
      Tag             =   "01-Country"
      Top             =   7920
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   15
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   95
      Tag             =   "01-Country"
      Top             =   8280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   16
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   96
      Tag             =   "01-Country"
      Top             =   8640
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   17
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   97
      Tag             =   "01-Country"
      Top             =   9000
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   18
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   98
      Tag             =   "01-Country"
      Top             =   9360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtCountry1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   19
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   99
      Tag             =   "01-Country"
      Top             =   9720
      Visible         =   0   'False
      Width           =   555
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   0
      Left            =   7440
      TabIndex        =   102
      Tag             =   "00-Currency Indicator - Code"
      Top             =   2880
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   0
      Left            =   10440
      TabIndex        =   103
      Tag             =   "21-Enter Exchange Rate"
      Top             =   2880
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   106
      Tag             =   "00-Currency Indicator - Code"
      Top             =   3240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   1
      Left            =   10440
      TabIndex        =   107
      Tag             =   "21-Enter Exchange Rate"
      Top             =   3240
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   2
      Left            =   7440
      TabIndex        =   110
      Tag             =   "00-Currency Indicator - Code"
      Top             =   3600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   2
      Left            =   10440
      TabIndex        =   111
      Tag             =   "21-Enter Exchange Rate"
      Top             =   3600
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   3
      Left            =   7440
      TabIndex        =   114
      Tag             =   "00-Currency Indicator - Code"
      Top             =   3960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   3
      Left            =   10440
      TabIndex        =   115
      Tag             =   "21-Enter Exchange Rate"
      Top             =   3960
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   4
      Left            =   7440
      TabIndex        =   118
      Tag             =   "00-Currency Indicator - Code"
      Top             =   4320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   4
      Left            =   10440
      TabIndex        =   119
      Tag             =   "21-Enter Exchange Rate"
      Top             =   4320
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   5
      Left            =   7440
      TabIndex        =   122
      Tag             =   "00-Currency Indicator - Code"
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   5
      Left            =   10440
      TabIndex        =   123
      Tag             =   "21-Enter Exchange Rate"
      Top             =   4680
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   6
      Left            =   7440
      TabIndex        =   126
      Tag             =   "00-Currency Indicator - Code"
      Top             =   5040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   6
      Left            =   10440
      TabIndex        =   127
      Tag             =   "21-Enter Exchange Rate"
      Top             =   5040
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   7
      Left            =   7440
      TabIndex        =   130
      Tag             =   "00-Currency Indicator - Code"
      Top             =   5400
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   7
      Left            =   10440
      TabIndex        =   131
      Tag             =   "21-Enter Exchange Rate"
      Top             =   5400
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   8
      Left            =   7440
      TabIndex        =   134
      Tag             =   "00-Currency Indicator - Code"
      Top             =   5760
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   8
      Left            =   10440
      TabIndex        =   135
      Tag             =   "21-Enter Exchange Rate"
      Top             =   5760
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   9
      Left            =   7440
      TabIndex        =   138
      Tag             =   "00-Currency Indicator - Code"
      Top             =   6120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   9
      Left            =   10440
      TabIndex        =   139
      Tag             =   "21-Enter Exchange Rate"
      Top             =   6120
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   10
      Left            =   7440
      TabIndex        =   142
      Tag             =   "00-Currency Indicator - Code"
      Top             =   6480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   10
      Left            =   10440
      TabIndex        =   143
      Tag             =   "21-Enter Exchange Rate"
      Top             =   6480
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   11
      Left            =   7440
      TabIndex        =   146
      Tag             =   "00-Currency Indicator - Code"
      Top             =   6840
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   11
      Left            =   10440
      TabIndex        =   147
      Tag             =   "21-Enter Exchange Rate"
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   12
      Left            =   7440
      TabIndex        =   150
      Tag             =   "00-Currency Indicator - Code"
      Top             =   7200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   12
      Left            =   10440
      TabIndex        =   151
      Tag             =   "21-Enter Exchange Rate"
      Top             =   7200
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   13
      Left            =   7440
      TabIndex        =   154
      Tag             =   "00-Currency Indicator - Code"
      Top             =   7560
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   13
      Left            =   10440
      TabIndex        =   155
      Tag             =   "21-Enter Exchange Rate"
      Top             =   7560
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   14
      Left            =   7440
      TabIndex        =   158
      Tag             =   "00-Currency Indicator - Code"
      Top             =   7920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   14
      Left            =   10440
      TabIndex        =   159
      Tag             =   "21-Enter Exchange Rate"
      Top             =   7920
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   15
      Left            =   7440
      TabIndex        =   162
      Tag             =   "00-Currency Indicator - Code"
      Top             =   8280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   15
      Left            =   10440
      TabIndex        =   163
      Tag             =   "21-Enter Exchange Rate"
      Top             =   8280
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   16
      Left            =   7440
      TabIndex        =   166
      Tag             =   "00-Currency Indicator - Code"
      Top             =   8640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   16
      Left            =   10440
      TabIndex        =   167
      Tag             =   "21-Enter Exchange Rate"
      Top             =   8640
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   17
      Left            =   7440
      TabIndex        =   170
      Tag             =   "00-Currency Indicator - Code"
      Top             =   9000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   17
      Left            =   10440
      TabIndex        =   171
      Tag             =   "21-Enter Exchange Rate"
      Top             =   9000
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   18
      Left            =   7440
      TabIndex        =   174
      Tag             =   "00-Currency Indicator - Code"
      Top             =   9360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   18
      Left            =   10440
      TabIndex        =   175
      Tag             =   "21-Enter Exchange Rate"
      Top             =   9360
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpCode2 
      Height          =   285
      Index           =   19
      Left            =   7440
      TabIndex        =   178
      Tag             =   "00-Currency Indicator - Code"
      Top             =   9720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "WFCI"
   End
   Begin MSMask.MaskEdBox medRate2 
      Height          =   285
      Index           =   19
      Left            =   10440
      TabIndex        =   179
      Tag             =   "21-Enter Exchange Rate"
      Top             =   9720
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "#,##0.00000;(#,##0.00000)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate"
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
      Index           =   8
      Left            =   10080
      TabIndex        =   80
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Indicator"
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
      Index           =   7
      Left            =   7440
      TabIndex        =   79
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Index           =   6
      Left            =   5880
      TabIndex        =   78
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate"
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
      Index           =   5
      Left            =   4320
      TabIndex        =   77
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Indicator"
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
      Index           =   4
      Left            =   1680
      TabIndex        =   76
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
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
      Index           =   3
      Left            =   120
      TabIndex        =   75
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year/Month"
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
      Index           =   2
      Left            =   120
      TabIndex        =   73
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   5760
      Y1              =   2160
      Y2              =   10080
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert to 2"
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
      Left            =   5880
      TabIndex        =   66
      Top             =   2325
      Width           =   1065
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert to 1"
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
      Left            =   120
      TabIndex        =   65
      Top             =   2325
      Width           =   1065
   End
End
Attribute VB_Name = "frmIPCurrencyExch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Actn
Dim fglbNew As Boolean
Dim SQLQ As String
Dim fglbESQLQ, fglbVSQLQ

Private Sub ComMTH_Change()
    txtMTH.Text = Left(ComMTH.Text, 2)
End Sub

Private Sub Form_Activate()
    glbOnTop = "frmIPCurrencyExch"
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    glbOnTop = "frmIPCurrencyExch"
    
    Call MonthDescAdd
    Call CountyAddItems
    
    data1.ConnectionString = glbAdoIHRDB
    
    SQLQ = "SELECT DISTINCT IP_YEAR,IP_MTH_SEQ,IP_MTH_DESC,IP_CURRENCYIND1,IP_CURRENCYIND2 FROM HRIP_CURRENCY_EXCHG "
    SQLQ = SQLQ & "ORDER BY IP_YEAR,IP_MTH_SEQ "
    data1.RecordSource = SQLQ
    data1.Refresh

    
    Call INI_Controls(Me)
    
End Sub

Private Sub MonthDescAdd()
ComMTH.AddItem "00-Annual Average Rate"
ComMTH.AddItem "01-January"
ComMTH.AddItem "02-February"
ComMTH.AddItem "03-March"
ComMTH.AddItem "04-April"
ComMTH.AddItem "05-May"
ComMTH.AddItem "06-June"
ComMTH.AddItem "07-July"
ComMTH.AddItem "08-August"
ComMTH.AddItem "09-September"
ComMTH.AddItem "10-October"
ComMTH.AddItem "11-November"
ComMTH.AddItem "12-December"
ComMTH.ListIndex = -1
End Sub

Private Sub CountyAddItems()
Dim I As Integer
Dim ctylist, x

ctylist = CountryList
x = 1
I = 0
Do While x > 0
    x = InStr(ctylist, "&")
    For I = 0 To 19
        If x > 0 Then
            'comCountry1(I).AddItem Left(ctylist, x - 1)
            'comCountry2(I).AddItem Left(ctylist, x - 1)
        Else
            'comCountry1(I).AddItem ctylist
            'comCountry2(I).AddItem ctylist
        End If
    Next
    ctylist = Mid(ctylist, x + 1)
Loop

For I = 0 To 19
    'comCountry1(I).ListIndex = -1
    'comCountry2(I).ListIndex = -1
Next

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

''If InStr(xCountryList, comCountry) = 0 And comCountry <> "" Then
''    xCountryList = xCountryList & "&" & comCountry
''    comCountry.AddItem comCountry
''    comCountryOfEmp.AddItem comCountry
''End If
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
RelateMode = nothingrelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Entitlements
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

Sub cmdNew_Click()

Call ResetAllFieldsToBlank

Actn = "A"

fglbNew = True

Call SET_UP_MODE

MskFiscalYear.SetFocus
End Sub

Sub ResetAllFieldsToBlank()
Dim I As Integer

MskFiscalYear.Text = ""
ComMTH.ListIndex = -1
clpCode(0).Text = ""
clpCode(1).Text = ""

For I = 0 To 19
    'comCountry1(I).Text = ""
    'comCountry1(I).ListIndex = -1
    'txtCountry1(I).Text = ""
    clpCode1(I).Text = ""
    medRate1(I).Text = ""
    
    'comCountry2(I).Text = ""
    'comCountry2(I).ListIndex = -1
    'txtCountry2(I).Text = ""
    clpCode2(I).Text = ""
    medRate2(I).Text = ""
Next

End Sub

Sub cmdCancel_Click()
fglbNew = False

data1.Refresh

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Call Display_Value

'orgEffDate = dlpAsOf.Text

vbxTrueGrid.SetFocus

End Sub

Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim x

Call ResetAllFieldsToBlank

If Not data1.Recordset.EOF Then
    SQLQ = "SELECT * FROM HRIP_CURRENCY_EXCHG WHERE (1=1) "
    SQLQ = SQLQ & " AND IP_YEAR = " & data1.Recordset("IP_YEAR") & " "
    SQLQ = SQLQ & " AND IP_MTH_SEQ = " & data1.Recordset("IP_MTH_SEQ") & " "
    'SQLQ = SQLQ & " AND IP_CURRENCYIND1 = '" & Data1.Recordset("IP_CURRENCYIND1") & "'"
    SQLQ = SQLQ & " Order By IP_YEAR,IP_MTH_SEQ,IP_CONVERT_NO,IP_ORDER "
    If rsVE.State <> 0 Then rsVE.Close
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsVE.EOF Then
        MskFiscalYear.Text = data1.Recordset("IP_YEAR")
        txtMTH.Text = data1.Recordset("IP_MTH_SEQ")
        ComMTH.ListIndex = getMonthIndex(data1.Recordset("IP_MTH_SEQ"))
        If Not IsNull(data1.Recordset("IP_CURRENCYIND1")) Then
            clpCode(0).Text = data1.Recordset("IP_CURRENCYIND1")
        End If
        If Not IsNull(data1.Recordset("IP_CURRENCYIND2")) Then
            clpCode(1).Text = data1.Recordset("IP_CURRENCYIND2")
        End If
    End If
    Do While Not rsVE.EOF
        xOrder = rsVE("IP_ORDER")
        If rsVE("IP_CONVERT_NO") = 1 Then 'convert 1
            'If Not IsNull(rsVE("IP_COUNTRY")) Then comCountry1(xOrder) = rsVE("IP_COUNTRY")
            If Not IsNull(rsVE("IP_CURRENCYINDF")) Then clpCode1(xOrder) = rsVE("IP_CURRENCYINDF")
            If Not IsNull(rsVE("IP_RATE")) Then medRate1(xOrder) = rsVE("IP_RATE")
        End If
        If rsVE("IP_CONVERT_NO") = 2 Then 'convert 2
            'If Not IsNull(rsVE("IP_COUNTRY")) Then comCountry2(xOrder) = rsVE("IP_COUNTRY")
            If Not IsNull(rsVE("IP_CURRENCYINDF")) Then clpCode2(xOrder) = rsVE("IP_CURRENCYINDF")
            If Not IsNull(rsVE("IP_RATE")) Then medRate2(xOrder) = rsVE("IP_RATE")
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close
End If

Call SET_UP_MODE
Call cmdModify_Click

End Sub

Sub cmdModify_Click()
'oDiv = MskFiscalYear.Text
'orgEffDate = dlpAsOf.Text
Actn = "M"
End Sub

Sub cmdOK_Click()
Dim x%, Y%, xUnion, xPT, SQLQ, SQLQW
Dim xStr
Dim rsVE As New ADODB.Recordset
Dim rsVT As New ADODB.Recordset
Dim glbiOneWhere As Boolean
Dim bmk As Variant

On Error GoTo AddN_Err

If data1.Recordset.EOF And data1.Recordset.BOF Then
    bmk = 0 'Ticket #11885 Frank Oct 11th, 2006
Else
    bmk = data1.Recordset.Bookmark
End If

If Not chkCurrency() Then Exit Sub

fglbVSQLQ = "IP_YEAR = " & MskFiscalYear.Text & " " ' data1.Recordset("IP_YEAR") & " "
fglbVSQLQ = fglbVSQLQ & " AND IP_MTH_SEQ = '" & Left(ComMTH.Text, 2) & "' "
    
If Actn = "M" Then
    'Call getWSQLQ("O")
    SQLQ = "DELETE FROM HRIP_CURRENCY_EXCHG WHERE " & fglbVSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    'Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HRIP_CURRENCY_EXCHG WHERE " & fglbVSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You can not add duplicate record"
        MskFiscalYear.SetFocus
        Exit Sub
    End If
End If

gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HRIP_CURRENCY_EXCHG "
rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ' , adOpenStatic, adLockOptimistic
For x% = 0 To 19
    If Len(clpCode1(x%)) > 0 Then 'convert 1
        rsVE.AddNew
        rsVE("IP_YEAR") = MskFiscalYear.Text
        rsVE("IP_MTH_SEQ") = Left(ComMTH.Text, 2)
        rsVE("IP_MTH_DESC") = Left(Mid(ComMTH.Text, 4, 30), 30)
        rsVE("IP_CURRENCYIND1") = clpCode(0).Text
        If Len(clpCode(1).Text) > 0 Then
            rsVE("IP_CURRENCYIND2") = clpCode(1).Text
        End If
        rsVE("IP_CONVERT_NO") = 1
        rsVE("IP_ORDER") = x%
        'If Len(comCountry1(x%).Text) > 0 Then
        '    rsVE("IP_COUNTRY") = Left(comCountry1(x%).Text, 10)
        'End If
        rsVE("IP_CURRENCYINDF") = clpCode1(x%).Text
        If Len(medRate1(x%).Text) > 0 Then
            If IsNumeric(medRate1(x%).Text) Then rsVE("IP_RATE") = medRate1(x%).Text
        End If
        rsVE("IP_LDATE") = Date
        rsVE("IP_LTIME") = Time$
        rsVE("IP_LUSER") = glbUserID
        rsVE.Update
    End If
    
    If Len(clpCode2(x%)) > 0 Then 'convert 2
        rsVE.AddNew
        rsVE("IP_YEAR") = MskFiscalYear.Text
        rsVE("IP_MTH_SEQ") = Left(ComMTH.Text, 2)
        rsVE("IP_MTH_DESC") = Left(Mid(ComMTH.Text, 4, 30), 30)
        rsVE("IP_CURRENCYIND1") = clpCode(0).Text
        If Len(clpCode(1).Text) > 0 Then
            rsVE("IP_CURRENCYIND2") = clpCode(1).Text
        End If
        rsVE("IP_CONVERT_NO") = 2
        rsVE("IP_ORDER") = x%
        'If Len(comCountry2(x%).Text) > 0 Then
        '    rsVE("IP_COUNTRY") = Left(comCountry2(x%).Text, 10)
        'End If
        rsVE("IP_CURRENCYINDF") = clpCode2(x%).Text
        If Len(medRate2(x%).Text) > 0 Then
            If IsNumeric(medRate2(x%).Text) Then rsVE("IP_RATE") = medRate2(x%).Text
        End If
        rsVE("IP_LDATE") = Date
        rsVE("IP_LTIME") = Time$
        rsVE("IP_LUSER") = glbUserID
        rsVE.Update
    End If
Next
rsVE.Close
gdbAdoIhr001.CommitTrans

Call Pause(0.5)
data1.Refresh

If Not bmk = 0 Then
    data1.Recordset.Bookmark = bmk
End If

fglbNew = False

Call Display_Value

'orgEffDate = dlpAsOf.Text

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

If Err.Number = -2147217887 Then '01/01/1200 can cause this error Ticket #18227
    MsgBox "    Invalid Date!    "
    gdbAdoIhr001.RollbackTrans
    Exit Sub
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "Currency Exchange Table", "UPDATE")
    Unload Me
End If
Resume Next

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Sub cmdDelete_Click()
Dim SQLQ, Msg, a%
If data1.Recordset.BOF And data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
Msg = "Are you sure you want to delete all records of " & Mid(ComMTH.Text, 4, 30) & " " & MskFiscalYear.Text & "?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

'Call getWSQLQ("C")
fglbVSQLQ = "IP_YEAR = " & MskFiscalYear.Text & " " ' data1.Recordset("IP_YEAR") & " "
fglbVSQLQ = fglbVSQLQ & " AND IP_MTH_SEQ = '" & Left(ComMTH.Text, 2) & "' "
SQLQ = "DELETE FROM HRIP_CURRENCY_EXCHG WHERE " & fglbVSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

data1.Refresh
Call Display_Value

'orgEffDate = dlpAsOf.Text

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%

'Me.vbxCrystal.Reset
'Me.vbxCrystal.WindowTitle = "Sick Entitlement Master Report"
'Call setRptLabel(Me, 0) '1)
'If glbSQL Or glbOracle Then
'    Me.vbxCrystal.Connect = RptODBC_SQL
'Else
'    Me.vbxCrystal.Connect = "PWD=petman;"
'    For x% = 0 To 5
'        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
'    Next
'End If
'Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgsickent.rpt"
'
'SQLQ = "(1=1) "
'If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_DIV} = '" & clpDiv.Text & "'"
'If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_DEPT} = '" & clpDept.Text & "'"
'If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_ORG} = '" & clpCode(0).Text & "'"
'If Len(clpCode(1).Text) > 0 Then SQLQ = SQLQ & " AND {HRSICKENT.VE_EMP} = '" & clpCode(1).Text & "'"

''sam 02/02/2006
'If Len(dlpDateRange(0).Text) > 0 Then
'    dtYYY% = Year(dlpDateRange(0).Text)
'    dtMM% = month(dlpDateRange(0).Text)
'    dtDD% = Day(dlpDateRange(0).Text)
'    SQLQ = SQLQ & " AND {HRSICKENT.VE_FRDATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'End If
'If Len(dlpDateRange(1).Text) > 0 Then
'    dtYYY% = Year(dlpDateRange(1).Text)
'    dtMM% = month(dlpDateRange(1).Text)
'    dtDD% = Day(dlpDateRange(1).Text)
'    SQLQ = SQLQ & " AND {HRSICKENT.VE_TODATE}  = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'End If
'
''sam 02/02/2006
'
'
'Me.vbxCrystal.SelectionFormula = SQLQ
'Me.vbxCrystal.Destination = 0
'Me.vbxCrystal.Action = 1

End Sub

Sub cmdPrint_Click()
    Call cmdView_Click
End Sub


Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
    'cmdPrintAll.Enabled = False
    'cmdUpdate.Enabled = False
    'CmdRecalc.Enabled = False
    'cmdUpdateAll.Enabled = False
ElseIf Me.data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    'cmdPrintAll.Enabled = True
    'cmdUpdate.Enabled = False
    'CmdRecalc.Enabled = False
    'cmdUpdateAll.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    'cmdPrintAll.Enabled = True
    'cmdUpdate.Enabled = True
    'CmdRecalc.Enabled = True
    'cmdUpdateAll.Enabled = True
End If

Call ST_UPD_MODE(TF)


Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

End Sub

Sub ST_UPD_MODE(TF As Boolean)
Dim I As Integer, FT

FT = Not TF

MskFiscalYear.Enabled = TF
ComMTH.Enabled = TF
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF

For I = 0 To 19
    comCountry1(I).Enabled = TF
    'txtCountry1(I).Enabled = TF
    clpCode1(I).Enabled = TF
    medRate1(I).Enabled = TF
    
    'comCountry2(I).Enabled = TF
    'txtCountry2(I).Enabled = TF
    clpCode2(I).Enabled = TF
    medRate2(I).Enabled = TF
Next

If data1.Recordset.EOF And data1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

'Call modSetFGlobals("SICK")

End Sub

Private Function getMonthIndex(xInx)
Dim retval As Integer

retval = -1
Select Case xInx
Case "00"
    retval = 0
Case "01"
    retval = 1
Case "02"
    retval = 2
    
Case "03"
    retval = 3
Case "04"
    retval = 4
Case "05"
    retval = 5
Case "06"
    retval = 6
Case "07"
    retval = 7
Case "08"
    retval = 8
Case "09"
    retval = 9
Case "10"
    retval = 10
Case "11"
    retval = 11
Case "12"
    retval = 12
End Select

getMonthIndex = retval
End Function

Private Function chkCurrency()
Dim x%, Y%

chkCurrency = False

On Error GoTo chkCurrency_Err

If Len(MskFiscalYear.Text) > 0 Then
    If Not IsNumeric(MskFiscalYear.Text) Then
        MsgBox "Invalid Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYear.Text) = 4 Then
        MsgBox "Invalid Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYear.SetFocus
    Exit Function
End If
If Len(ComMTH.Text) = 0 Then
    MsgBox "Month is a required field"
    ComMTH.SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) = 0 Then
    MsgBox "Convert to 1 is a required field"
    clpCode(0).SetFocus
    Exit Function
End If

For x% = 0 To 1
    If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
        MsgBox "If Code entered it must be known"
        clpCode(x%).SetFocus
        Exit Function
    End If
Next x%
For x% = 0 To 19
    If Len(clpCode1(x%).Text) > 0 And clpCode1(x%).Caption = "Unassigned" Then
        MsgBox "If Code entered it must be known"
        clpCode1(x%).SetFocus
        Exit Function
    End If
    If Len(clpCode2(x%).Text) > 0 And clpCode2(x%).Caption = "Unassigned" Then
        MsgBox "If Code entered it must be known"
        clpCode2(x%).SetFocus
        Exit Function
    End If
Next x%



If Len(clpCode1(0)) < 1 Then
    MsgBox "You must have at least one Currency Entry."
    If clpCode1(0).Enabled Then clpCode1(0).SetFocus
    Exit Function
End If
If Len(clpCode2(0)) > 0 Then
    If Len(clpCode(1).Text) = 0 Then
        MsgBox "Convert to 2 Code is a required field if there is Currency entered under Convert to 2 list"
        If clpCode2(0).Enabled Then clpCode2(0).SetFocus
        Exit Function
    End If
End If

For x% = 0 To 19
    If Len(medRate1(x%)) > 0 Then
        If Not IsNumeric(medRate1(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medRate1(x%).SetFocus
            Exit Function
        End If
    End If
    If Len(medRate2(x%)) > 0 Then
        If Not IsNumeric(medRate2(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medRate2(x%).SetFocus
            Exit Function
        End If
    End If

    'If Len(medLTServ(x%)) < 1 Then Exit For  ' missed one
    'intRangesSet% = intRangesSet% + 1
Next x%


chkCurrency = True

Exit Function

chkCurrency_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkCurrency", "HRIP_CURRENCY_EXCHG", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub
