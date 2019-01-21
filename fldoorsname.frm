VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSDoorsName 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Door Name Setup"
   ClientHeight    =   8595
   ClientLeft      =   525
   ClientTop       =   1470
   ClientWidth     =   8880
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
   ScaleHeight     =   8595
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   4785
      LargeChange     =   315
      Left            =   8280
      Max             =   100
      SmallChange     =   315
      TabIndex        =   55
      Top             =   2040
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   33
      Top             =   7935
      Width           =   8880
      _Version        =   65536
      _ExtentX        =   15663
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
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Tag             =   "Print All Records"
         Top             =   120
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   6960
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
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fldoorsname.frx":0000
      Height          =   1845
      Left            =   0
      OleObjectBlob   =   "fldoorsname.frx":0014
      TabIndex        =   0
      Top             =   150
      Width           =   8055
   End
   Begin VB.Frame frmShow 
      BorderStyle     =   0  'None
      Height          =   6075
      Left            =   120
      TabIndex        =   34
      Top             =   2040
      Width           =   7995
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Doors"
         Height          =   375
         Index           =   4
         Left            =   6060
         TabIndex        =   75
         Tag             =   "Print Door Access"
         Top             =   4890
         Width           =   1395
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Doors"
         Height          =   375
         Index           =   3
         Left            =   6060
         TabIndex        =   74
         Tag             =   "Print Door Access"
         Top             =   3780
         Width           =   1395
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Doors"
         Height          =   375
         Index           =   2
         Left            =   6060
         TabIndex        =   73
         Tag             =   "Print Door Access"
         Top             =   2640
         Width           =   1395
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Doors"
         Height          =   375
         Index           =   1
         Left            =   6060
         TabIndex        =   72
         Tag             =   "Print Door Access"
         Top             =   1530
         Width           =   1395
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Doors"
         Height          =   375
         Index           =   0
         Left            =   6060
         TabIndex        =   71
         Tag             =   "Print Door Access"
         Top             =   390
         Width           =   1395
      End
      Begin VB.Frame frmCtrl 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   3
         Left            =   1620
         TabIndex        =   68
         Top             =   5040
         Width           =   1815
         Begin VB.OptionButton opt9560 
            Caption         =   "9560"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   795
         End
         Begin VB.OptionButton opt2480 
            Caption         =   "2480"
            Height          =   255
            Index           =   4
            Left            =   900
            TabIndex        =   27
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame frmCtrl 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   2
         Left            =   1620
         TabIndex        =   66
         Top             =   3930
         Width           =   1815
         Begin VB.OptionButton opt9560 
            Caption         =   "9560"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   795
         End
         Begin VB.OptionButton opt2480 
            Caption         =   "2480"
            Height          =   255
            Index           =   3
            Left            =   900
            TabIndex        =   21
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame frmCtrl 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   1
         Left            =   1620
         TabIndex        =   64
         Top             =   2790
         Width           =   1815
         Begin VB.OptionButton opt2480 
            Caption         =   "2480"
            Height          =   255
            Index           =   2
            Left            =   900
            TabIndex        =   15
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton opt9560 
            Caption         =   "9560"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.Frame frmCtrl 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   4
         Left            =   1620
         TabIndex        =   61
         Top             =   1650
         Width           =   1815
         Begin VB.OptionButton opt9560 
            Caption         =   "9560"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   795
         End
         Begin VB.OptionButton opt2480 
            Caption         =   "2480"
            Height          =   255
            Index           =   1
            Left            =   900
            TabIndex        =   9
            Top             =   0
            Width           =   915
         End
      End
      Begin VB.Frame frmCtrl 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   60
         Top             =   510
         Width           =   1815
         Begin VB.OptionButton opt2480 
            Caption         =   "2480"
            Height          =   255
            Index           =   0
            Left            =   900
            TabIndex        =   3
            Top             =   0
            Width           =   915
         End
         Begin VB.OptionButton opt9560 
            Caption         =   "9560"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.OptionButton opt2480 
         Caption         =   "2480"
         Height          =   255
         Index           =   21
         Left            =   5100
         TabIndex        =   59
         Top             =   8400
         Width           =   915
      End
      Begin VB.OptionButton opt9560 
         Caption         =   "9560"
         Height          =   255
         Index           =   21
         Left            =   4200
         TabIndex        =   58
         Top             =   8400
         Width           =   915
      End
      Begin VB.OptionButton opt2480 
         Caption         =   "2480"
         Height          =   255
         Index           =   20
         Left            =   5100
         TabIndex        =   57
         Top             =   7980
         Width           =   915
      End
      Begin VB.OptionButton opt9560 
         Caption         =   "9560"
         Height          =   255
         Index           =   20
         Left            =   4200
         TabIndex        =   56
         Top             =   7980
         Width           =   915
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME1"
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
         Left            =   705
         MaxLength       =   50
         TabIndex        =   4
         Top             =   810
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME2"
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
         Left            =   705
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1140
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME3"
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
         Left            =   4545
         MaxLength       =   50
         TabIndex        =   6
         Top             =   810
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME4"
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
         Index           =   3
         Left            =   4545
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1140
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME5"
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
         Index           =   4
         Left            =   705
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1980
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME6"
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
         Index           =   5
         Left            =   705
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2310
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME7"
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
         Index           =   6
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1950
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME8"
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
         Index           =   7
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2280
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME9"
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
         Index           =   8
         Left            =   705
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3120
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME10"
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
         Index           =   9
         Left            =   705
         MaxLength       =   50
         TabIndex        =   17
         Top             =   3450
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME11"
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
         Index           =   10
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3060
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME12"
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
         Index           =   11
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   19
         Top             =   3450
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME13"
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
         Index           =   12
         Left            =   705
         MaxLength       =   50
         TabIndex        =   22
         Top             =   4230
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME14"
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
         Index           =   13
         Left            =   705
         MaxLength       =   50
         TabIndex        =   23
         Top             =   4560
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME15"
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
         Index           =   14
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   24
         Top             =   4230
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME16"
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
         Index           =   15
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   25
         Top             =   4560
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME17"
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
         Index           =   16
         Left            =   705
         MaxLength       =   50
         TabIndex        =   28
         Top             =   5310
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME18"
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
         Index           =   17
         Left            =   705
         MaxLength       =   50
         TabIndex        =   29
         Top             =   5640
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME19"
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
         Index           =   18
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   30
         Top             =   5310
         Width           =   3000
      End
      Begin VB.TextBox txtDoorName 
         Appearance      =   0  'Flat
         DataField       =   "DOORNAME20"
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
         Index           =   19
         Left            =   4605
         MaxLength       =   50
         TabIndex        =   31
         Top             =   5640
         Width           =   3000
      End
      Begin INFOHR_Controls.CodeLookup clpDIV 
         DataField       =   "DIV"
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Tag             =   "01-Division"
         Top             =   30
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
         Object.Height          =   315
      End
      Begin VB.Label lblTitle 
         Caption         =   "Facility"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   70
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Door Controller 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   69
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Door Controller 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   67
         Top             =   3930
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Door Controller 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   65
         Top             =   2790
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Door Controller 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   1650
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Door Controller 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   62
         Top             =   510
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 1"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 2"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   53
         Top             =   1170
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 3"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   52
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 4"
         Height          =   255
         Index           =   4
         Left            =   3840
         TabIndex        =   51
         Top             =   1170
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 5"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   50
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 6"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   49
         Top             =   2340
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 7"
         Height          =   255
         Index           =   7
         Left            =   3900
         TabIndex        =   48
         Top             =   1980
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 8"
         Height          =   255
         Index           =   8
         Left            =   3900
         TabIndex        =   47
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 9"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   46
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 10"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   45
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 11"
         Height          =   255
         Index           =   11
         Left            =   3900
         TabIndex        =   44
         Top             =   3090
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 12"
         Height          =   255
         Index           =   12
         Left            =   3900
         TabIndex        =   43
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 13"
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   42
         Top             =   4260
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 14"
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   41
         Top             =   4590
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 15"
         Height          =   255
         Index           =   15
         Left            =   3900
         TabIndex        =   40
         Top             =   4260
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 16"
         Height          =   255
         Index           =   16
         Left            =   3900
         TabIndex        =   39
         Top             =   4590
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 17"
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   38
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 18"
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   37
         Top             =   5670
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 19"
         Height          =   255
         Index           =   19
         Left            =   3900
         TabIndex        =   36
         Top             =   5340
         Width           =   1035
      End
      Begin VB.Label lblTitle 
         Caption         =   "Door 20"
         Height          =   255
         Index           =   20
         Left            =   3900
         TabIndex        =   35
         Top             =   5670
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmSDoorsName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNewRec%
Dim DCtl(20)
Dim rsDATA As New ADODB.Recordset
Dim fglbNew
Sub cmdCancel_Click()

On Error GoTo Can_Err

rsDATA.CancelUpdate
Data1.Refresh
Call ST_UPD_MODE(False)  ' reset screen's attributes
Me.vbxTrueGrid.SetFocus

fglbNewRec% = False
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



 Sub cmdClose_Click()
Unload Me

End Sub





Sub cmdOK_Click()
Dim X%
Dim xID
Dim i, j
On Error GoTo OK_Err

If Not chkDoorName Then Exit Sub

For i = 1 To 20
    j = Int((i - 1) / 4)
    rsDATA("DOORCTRL" & i) = IIf(opt9560(j).Value, "9560", IIf(opt2480(j).Value, "2480", Null))
Next
Call Set_Control("U", Me, rsDATA)
gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
fglbNew = False
Call ST_UPD_MODE(False)


fglbNewRec% = False
Me.vbxTrueGrid.SetFocus

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



Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%



Me.vbxCrystal.WindowTitle = "Door Names Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "Rldoornm.rpt"
    Me.vbxCrystal.SelectionFormula = "{LN_DOORS_NAME.DIV}='" & clpDIV & "' "
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1


End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, X%



Me.vbxCrystal.WindowTitle = "Door Names Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "Rldoornm.rpt"
    Me.vbxCrystal.SelectionFormula = "{LN_DOORS_NAME.DIV}='" & clpDIV & "' "
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1



End Sub
Private Sub cmdPrintAll_Click()
Dim RHeading As String, xReport, X%

cmdPrintAll.Enabled = False


Me.vbxCrystal.WindowTitle = "Door Names Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "Rldoornm.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True


End Sub

Private Sub cmdPrintAll_GotFocus()
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
Dim X%
glbOnTop = Me.name
Screen.MousePointer = HOURGLASS
frmSDoorsName.Show
Me.Caption = lStr("Door Name Setup - ")
Call INIData
Call EERetrieve
If vbxTrueGrid.Visible Then Me.vbxTrueGrid.SetFocus
Call INI_Controls(Me)

Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
If Me.Height >= 2100 + frmShow.Height + panControls.Height + 230 Then
    scrControl.Value = 0
    frmShow.Top = 2100
    scrControl.Visible = False
    Exit Sub
End If
scrControl.Visible = True
scrControl.Max = frmShow.Height + panControls.Height + 2100 + 550 - Me.Height '250 - Me.Height
scrControl.Left = Me.Width - scrControl.Width - 120
If Me.Height - scrControl.Top - panControls.Height - 300 > 0 Then
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 300
Else
    scrControl.Height = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmLDoors = Nothing

End Sub

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
Call ST_UPD_MODE(TF)
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
frmShow.Enabled = TF
End Sub







Private Sub cmdUpdate_Click(Index As Integer)
Dim xPath, xCtl, xNoBadgeID, xFile
Dim xDiv, xTitle, xPP, FirstDoor
Dim xDoors, i, j, k
Dim HaveDoorName
Dim rsDR As New ADODB.Recordset
xPath = App.Path
xPath = xPath & IIf(Right(xPath, 1) = "/", "", "/")


If Data1.Recordset.EOF Then
    MsgBox "Nothing to update."
    Exit Sub
End If

xCtl = ""
If opt9560(Index) Then xCtl = "9560"
If opt2480(Index) Then xCtl = "2480"
If xCtl = "" Then
    MsgBox "Door Controller must to be setup."
    Exit Sub
End If

FirstDoor = (Index * 4 + 1)
HaveDoorName = False
For j = 0 To 3
    k = FirstDoor + j
    If Trim(txtDoorName(k - 1)) <> "" Then
        HaveDoorName = True
    End If
Next
If Not HaveDoorName Then
    MsgBox "No Door Name was given."
    Exit Sub
End If

xDiv = Data1.Recordset!Div
rsDR.Open "SELECT * FROM LN_DOORS WHERE DIV='" & xDiv & "'", gdbAdoIhr001, adOpenForwardOnly

If rsDR.EOF Then
    MsgBox "No User/Employee access the doors."
    Exit Sub
End If

xFile = xPath & xDiv & "_" & xCtl & ".txt"
Open xFile For Output As #1

xNoBadgeID = xPath & xDiv & "_" & xCtl & "_NoBadgeID.txt"
Open xNoBadgeID For Output As #3

xTitle = "DU(0)=""" & (Index * 4 + 1) & """" & vbNewLine
xTitle = xTitle & "DU(1)=""" & Weekday(Date) & """" & vbNewLine
xTitle = xTitle & "DU(2)=""" & Format(Now, "yy/mm/dd:hh:mm:ss") & """" & vbNewLine
xTitle = xTitle & "DU(3)=""" & Data1.Recordset!division_name & """" '& vbNewLine
Print #1, xTitle
i = 1
Do Until rsDR.EOF
    If IsNull(rsDR!badgeid) Then
        Print #3, IIf(rsDR!EMP, "Employee #", "User ID") & rsDR!USERID '& vbTab & lblEEName
    Else
        xDoors = ""
        For j = 0 To 3
            k = FirstDoor + j
            xDoors = xDoors & "," & IIf(rsDR("DOOR" & k), k, " ")
        Next
        Print #1, "DA(" & i & ")=""" & rsDR!badgeid & vbTab & xDoors & """" '& vbNewLine
    End If
    rsDR.MoveNext
    i = i + 1
Loop
Print #1, ""
Print #3, ""
Close #1
Close #3
MsgBox "Finished"
End Sub

Private Sub opt2480_Click(Index As Integer)
Dim j
j = (Index + 1) * 4 - 1
If opt2480(Index) Then
    txtDoorName(j).Enabled = txtDoorName(j - 1).Enabled
End If
End Sub

Private Sub opt9560_Click(Index As Integer)
Dim j
j = (Index + 1) * 4 - 1
If opt9560(Index) Then
    txtDoorName(j) = ""
    txtDoorName(j).Enabled = False
End If
End Sub

Private Sub scrControl_Change()
frmShow.Top = 2100 - scrControl.Value
End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
         SQLQ = "select LN_DOORS_NAME.*,division_name from LN_DOORS_NAME inner join hr_division on hr_division.div =ln_doors_name.div where ln_doors_name.div in (select div from hr_division where " & glbSeleDiv & ") "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub



Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
End If

End Sub




Private Function chkDoorName()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset

chkDoorName = False
On Error GoTo chkDoorName_Err

If Len(clpDIV) < 1 Then
    MsgBox lStr("Division Code is a required field")
    clpDIV.SetFocus
    Exit Function
End If


If glbLinamar And (Len(clpDIV) <> 3 Or Not IsNumeric(clpDIV)) Then
    MsgBox lStr("Invalid Division")
    If clpDIV.Enabled Then clpDIV.SetFocus
    Exit Function
End If
If fglbNewRec Then
    Div = CStr(clpDIV)
    SQLQ = "SELECT DIV from LN_DOORS "
    SQLQ = SQLQ & "WHERE DIV = '" & Div & "'"
    
    If snapDivs.State <> 0 Then snapDivs.Close
    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDivs.BOF And snapDivs.EOF Then
        snapDivs.Close
    Else
        Msg$ = lStr("This Division number already exists")
        MsgBox Msg$
        snapDivs.Close
        Exit Function
    End If
End If

chkDoorName = True

Exit Function

chkDoorName_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDoorName", "LN_Doors_name", "Cancel")
Resume Next

End Function

Private Sub INIData()
Dim rsTD As New ADODB.Recordset
Dim rsSR As New ADODB.Recordset
Dim SQLQ
rsTD.Open "SELECT * FROM HR_DIVISION WHERE DIV<>'ALL'", gdbAdoIhr001, adOpenStatic
Do Until rsTD.EOF
    rsSR.Open "SELECT * FROM LN_DOORS_NAME WHERE DIV='" & rsTD("DIV") & "'", gdbAdoIhr001, adOpenStatic
    If rsSR.EOF Then
        SQLQ = "INSERT INTO LN_DOORS_NAME(DIV) VALUES('" & rsTD("DIV") & "')"
        gdbAdoIhr001.Execute SQLQ
    End If
    rsSR.Close
    rsTD.MoveNext
Loop
rsTD.Close
'Data1.Refresh
End Sub





Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub
Public Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "select LN_DOORS_NAME.*,division_name from LN_DOORS_NAME inner join hr_division on hr_division.div =ln_doors_name.div where ln_doors_name.div in (select div from hr_division where " & glbSeleDiv & ") ORDER BY division_name"
Data1.Refresh
EERetrieve = True
Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack '21June99 js


End Function


Sub Display_Value()
Dim SQLQ
Dim i
Dim j
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    If glbtermopen Then
        SQLQ = "Select Term_DOLENT.* from Term_DOLENT"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select HRDOLENT.* from HRDOLENT"
        SQLQ = SQLQ & " where DE_EMPNBR = " & glbLEE_ID
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
      
    End If
    Call SET_UP_MODE
'    Me.cmdModify_Click
    Exit Sub
End If
    
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close

SQLQ = "select * from LN_DOORS_NAME where ID=" & Data1.Recordset!ID
rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
For i = 1 To 5
    j = (i - 1) * 4 + 1
    DCtl(i - 1) = rsDATA("DOORCTRL" & j)
    opt9560(i - 1).Value = IIf(DCtl(i - 1) = "9560", 1, 0)
    opt2480(i - 1).Value = IIf(DCtl(i - 1) = "2480", 1, 0)
Next
Call SET_UP_MODE
'Me.cmdModify_Click
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
RelateMode = NothingRelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Security
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
Printable = True
End Property
