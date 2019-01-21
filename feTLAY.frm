VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmETLAY 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Temporary Lay Off"
   ClientHeight    =   9240
   ClientLeft      =   315
   ClientTop       =   780
   ClientWidth     =   10845
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
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9240
   ScaleWidth      =   10845
   WindowState     =   2  'Maximized
   Begin VB.Frame frmWFCBenList 
      Height          =   1215
      Left            =   240
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   9015
      Begin VB.CheckBox chkAllDates 
         Caption         =   "All Date"
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
         Left            =   4560
         TabIndex        =   49
         Top             =   2445
         Width           =   1155
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid1 
         Bindings        =   "feTLAY.frx":0000
         Height          =   2025
         Left            =   120
         OleObjectBlob   =   "feTLAY.frx":0014
         TabIndex        =   20
         Top             =   240
         Width           =   10275
      End
      Begin INFOHR_Controls.DateLookup dlpEndDate 
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1800
         TabIndex        =   50
         Tag             =   "41-Effective date of salary change"
         Top             =   2400
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Left            =   360
         TabIndex        =   51
         Top             =   2460
         Width           =   885
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   9240
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin Threed.SSPanel Panel3D1 
      Height          =   8220
      Left            =   240
      TabIndex        =   28
      Top             =   600
      Width           =   10035
      _Version        =   65536
      _ExtentX        =   17701
      _ExtentY        =   14499
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
      Begin VB.TextBox memComments 
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
         Height          =   1455
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Tag             =   "00-Comments"
         Top             =   6420
         Visible         =   0   'False
         Width           =   9045
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   270
         Left            =   8955
         TabIndex        =   16
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox comESalInc 
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
         Left            =   3260
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "Eligible for Salary Increase"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkATPaidHours 
         Caption         =   "Paid Hours in AT"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Frame frmMulti 
         Caption         =   "Position/Salary Information"
         Height          =   2865
         Left            =   5520
         TabIndex        =   35
         Top             =   120
         Width           =   4275
         Begin VB.TextBox txtPayrollID 
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
            Left            =   1930
            MaxLength       =   25
            TabIndex        =   15
            Tag             =   "00-Payroll ID"
            Top             =   2500
            Width           =   1815
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
            Left            =   1935
            TabIndex        =   11
            Tag             =   "Choose annum or hour"
            Top             =   1212
            Width           =   1215
         End
         Begin VB.TextBox txtWHRS 
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
            Left            =   1935
            TabIndex        =   13
            Tag             =   "00- Number of hours in work week"
            Top             =   1870
            Width           =   975
         End
         Begin VB.TextBox txtDHRS 
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
            Left            =   1935
            TabIndex        =   12
            Tag             =   "00-Usual working hours per day"
            Top             =   1556
            Width           =   855
         End
         Begin VB.CommandButton cmdPostion 
            Caption         =   "P&ositions"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Tag             =   "Postions"
            Top             =   285
            Width           =   975
         End
         Begin INFOHR_Controls.CodeLookup clpJob 
            Height          =   285
            Left            =   1620
            TabIndex        =   8
            Tag             =   "01-Job Code"
            Top             =   270
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   503
            ShowUnassigned  =   1
            ShowDescription =   0   'False
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   5
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   0
            Left            =   1620
            TabIndex        =   9
            Tag             =   "00-Union Code"
            Top             =   584
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin MSMask.MaskEdBox medSalary 
            Height          =   285
            Left            =   1935
            TabIndex        =   10
            Tag             =   "00-Usual working Salary"
            Top             =   898
            Width           =   1305
            _ExtentX        =   2302
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
         Begin INFOHR_Controls.CodeLookup clpGLNo 
            Height          =   285
            Left            =   1620
            TabIndex        =   14
            Tag             =   "00-General Ledger - Code"
            Top             =   2184
            Width           =   2505
            _ExtentX        =   4419
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   3
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "G/L #"
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
            Index           =   20
            Left            =   180
            TabIndex        =   44
            Top             =   2229
            Width           =   435
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
            Left            =   180
            TabIndex        =   43
            Top             =   2545
            Width           =   675
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Position"
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
            Left            =   180
            TabIndex        =   42
            Top             =   330
            Width           =   780
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Union"
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
            Left            =   180
            TabIndex        =   41
            Top             =   629
            Width           =   660
         End
         Begin VB.Label lblSalCode 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "H/A"
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
            Height          =   195
            Left            =   2790
            TabIndex        =   40
            Top             =   1260
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblHrsDay 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Day"
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
            Left            =   180
            TabIndex        =   39
            Top             =   1601
            Width           =   855
         End
         Begin VB.Label lblHrsWeek 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Hours/Week"
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
            Left            =   180
            TabIndex        =   38
            Top             =   1915
            Width           =   1095
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
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
            Index           =   10
            Left            =   180
            TabIndex        =   37
            Top             =   943
            Width           =   540
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Per"
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
            Left            =   180
            TabIndex        =   36
            Top             =   1272
            Width           =   300
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   2940
         TabIndex        =   3
         Tag             =   "41-Employement Status Code"
         Top             =   1800
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   2940
         TabIndex        =   0
         Tag             =   "41-Termination Code "
         Top             =   600
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin INFOHR_Controls.DateLookup dlpTLAYDate 
         Height          =   285
         Index           =   1
         Left            =   2940
         TabIndex        =   2
         Tag             =   "41-To Date"
         Top             =   1410
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpTLAYDate 
         Height          =   285
         Index           =   0
         Left            =   2940
         TabIndex        =   1
         Tag             =   "40-From Date"
         Top             =   1020
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.CheckBox chkLeave 
         Caption         =   "Leave"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Visible         =   0   'False
         Width           =   1515
      End
      Begin INFOHR_Controls.CodeLookup clpATTCode 
         Height          =   285
         Left            =   2940
         TabIndex        =   4
         Top             =   2190
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin INFOHR_Controls.DateLookup dlpDOther2 
         DataSource      =   " "
         Height          =   285
         Left            =   2940
         TabIndex        =   17
         Tag             =   "40-Other Date 2"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1045
      End
      Begin INFOHR_Controls.DateLookup dlpLastDat2 
         Height          =   285
         Left            =   6720
         TabIndex        =   19
         Tag             =   "41-Effective date of salary change"
         Top             =   3600
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Image imgSec 
         Height          =   240
         Left            =   8535
         Picture         =   "feTLAY.frx":4C49
         Top             =   3120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblImport 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "LOA"
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
         Height          =   240
         Left            =   7920
         TabIndex        =   54
         Top             =   3120
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image imgNoSec 
         Height          =   240
         Left            =   8535
         Picture         =   "feTLAY.frx":4D93
         Top             =   3120
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblComment 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   360
         TabIndex        =   53
         Top             =   6120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Day"
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
         Index           =   7
         Left            =   5280
         TabIndex        =   52
         Top             =   3660
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Eligible for Salary Increase"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   47
         Top             =   3990
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lbOtherDate2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Date 2"
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
         Left            =   360
         TabIndex        =   46
         Top             =   3600
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "Attendance Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   34
         Top             =   2220
         Width           =   1995
      End
      Begin VB.Label lblNotice 
         Caption         =   "This employee has been temporarily laid off for the dates shown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   420
         TabIndex        =   33
         Top             =   3180
         Visible         =   0   'False
         Width           =   7275
      End
      Begin VB.Label lblWeeks 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 Week"
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
         Left            =   4680
         TabIndex        =   32
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Employment Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   31
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "                               To"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   30
         Tag             =   "41-Date Terminated"
         Top             =   1410
         Width           =   2100
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date                        From"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Tag             =   "41-Date Terminated"
         Top             =   1020
         Width           =   2265
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Temporary Lay Off Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   27
         Top             =   570
         Width           =   2280
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   10845
      _Version        =   65536
      _ExtentX        =   19129
      _ExtentY        =   873
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
         TabIndex        =   45
         Top             =   127
         Width           =   1305
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
         Caption         =   "Employee"
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
         TabIndex        =   29
         Top             =   127
         Width           =   1065
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   150
         Width           =   1005
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   4440
         TabIndex        =   25
         Top             =   150
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   2880
         TabIndex        =   24
         Top             =   127
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   8760
      Top             =   5040
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   9240
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
Attribute VB_Name = "frmETLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbFollowID
Dim fglbNew
Dim OLDEMP
Dim oFDate, OTDate
Dim fdFdate, fdTdate
Dim fglbJobList As String
Dim xlocUSA As String
Dim MailBody As String
Dim xLocID '07/16/2013
Dim AbortLeave As Boolean
Dim xSamuleFlag As Boolean
Dim xNewStatus 'Ticket #30446 Franks 08/09/2017
Dim fglbWDate$

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim dynaJobHIS As New ADODB.Recordset
On Error GoTo JobHis_Err
fglbJobList = ""
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

Private Function AUDITTERM()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String, XSNAME As String, XFNAME As String, xEmpType As String
Dim SQLQ
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITTERM = False
xlocUSA = ""

rsTB.Open "SELECT ED_EMPNBR,ED_PT,ED_DIV,ED_SURNAME,ED_FNAME,ED_EMPTYPE,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & lblEEID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    
    'Ticket #30054 - Getting invalid use of null error when ED_PT is null
    If Not IsNull(rsTB("ED_PT")) Then
        xPT = rsTB("ED_PT")
    Else
        xPT = ""
    End If
    If Not IsNull(rsTB("ED_DIV")) Then
        xDiv = rsTB("ED_DIV")
    Else
        xDiv = ""
    End If
    XSNAME = rsTB("ED_SURNAME")
    XFNAME = rsTB("ED_FNAME")
    xEmpType = IIf(IsNull(rsTB("ED_EMPTYPE")), "", rsTB("ED_EMPTYPE"))
    OLDEMP = rsTB("ED_EMP")
Else
    xPT = ""
    xDiv = ""
    XSNAME = ""
    XFNAME = ""
    xEmpType = ""
    OLDEMP = ""
End If

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_EMP, AU_USRDAT1, AU_UNION, AU_SFDATE, "
strFields = strFields & "AU_STDATE, AU_SURNAME, AU_FNAME, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, "
strFields = strFields & "AU_TYPE, AU_PAYROLL_ID,AU_PENSION,AU_LTHIRE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_EMPTYPE") = xEmpType
rsTA("AU_EMP") = clpCode(2).Text

If glbLinamar Then
    rsTA("AU_USRDAT1") = dlpTLAYDate(0).Text
    rsTA("AU_UNION") = dlpTLAYDate(1).Text
Else
    rsTA("AU_SFDATE") = dlpTLAYDate(0).Text
    If Len(dlpTLAYDate(1).Text) > 0 Then 'Ticket #22009 Franks 05/10/2012
        rsTA("AU_STDATE") = dlpTLAYDate(1).Text
    End If
End If
'Ticket #18306 remove this from interface
'If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18267
'    rsTA("AU_LTHIRE") = dlpTLAYDate(0).Text
'End If

'rsTA("AU_SURNAME") = XSNAME
'rsTA("AU_FNAME") = XFNAME
'rsTA("AU_DOT") = dlpTLAYDate(0)
'rsTA("AU_TREAS") = clpCode(1)

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = lblEEID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"

'rsTA("AU_TYPE") = "T"
'Ticket #21181 Franks 11/09/2011 - use AU_TYPE as L for ther Audit report
'rsTA("AU_TYPE") = "M"
rsTA("AU_TYPE") = "L"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    'Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_WORKCOUNTRY FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        'Ticket #16749 for US ADP Payforce: L - Enter Leave
        If glbWFC Then
            If Not IsNull(rsEmp("ED_WORKCOUNTRY")) Then
                If rsEmp("ED_WORKCOUNTRY") = "U.S.A." Then
                    rsTA("AU_TYPE") = "L"
                    xlocUSA = "Y"
                    rsTA("AU_PENSION") = "2"
                End If
            End If
        End If
    End If
    rsEmp.Close
'End If

rsTA.Update

AUDITTERM = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js

End Function

Private Function chkTerms()
Dim dd As Integer
Dim X
Dim Msg$, DgDef As Variant, Response%, Title$ 'Ticket #24061 Franks 07/16/2013
Dim SQLQ As String 'Ticket #24094 Franks 07/22/2013
Dim rsTemp As New ADODB.Recordset

chkTerms = False
If Len(dlpTLAYDate(0).Text) < 1 Then
    MsgBox "From Date is a required field"
    dlpTLAYDate(0).SetFocus
    Exit Function
End If
If glbWFC Then 'Ticket #22009 Franks 05/10/2012
    If Len(clpATTCode.Text) > 0 Then
        If Len(dlpTLAYDate(1).Text) < 1 Then
            MsgBox "To Date is a required field if Attendance Code is enterred."
            dlpTLAYDate(1).SetFocus
            Exit Function
        End If
    End If
Else
    If Len(dlpTLAYDate(1).Text) < 1 Then
        MsgBox "To Date is a required field"
        dlpTLAYDate(1).SetFocus
        Exit Function
    End If
End If
If Not IsDate(dlpTLAYDate(0).Text) Then
    MsgBox "From Date is not a valid date"
    dlpTLAYDate(0).SetFocus
    Exit Function
End If

If Len(dlpTLAYDate(1).Text) > 0 Then
    If Not IsDate(dlpTLAYDate(1).Text) Then
        MsgBox "To Date is not a valid date"
        dlpTLAYDate(1).SetFocus
        Exit Function
    End If
End If

If Len(dlpTLAYDate(1).Text) > 0 Then
    If DateDiff("d", dlpTLAYDate(1), dlpTLAYDate(0)) > 0 Then
        MsgBox "From date must be earlier than To Date"
        dlpTLAYDate(0).SetFocus
        Exit Function
    End If
End If

' If statement above should work, but in any case I add If statement under afther test result
If IsDate(dlpTLAYDate(0).Text) And IsDate(dlpTLAYDate(1).Text) Then
    If DaysBetween(dlpTLAYDate(0), dlpTLAYDate(1)) < 0 Then
        MsgBox "From Date can not be prior to To Date"
        dlpTLAYDate(0).SetFocus
        Exit Function
    End If
End If

If IsDate(dlpTLAYDate(0).Text) Then
Dim rsEM As New ADODB.Recordset
rsEM.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If IsNull(rsEM("ED_DOH")) Then
        MsgBox "Original Hire Date cannot be blank"
        rsEM.Close
        dlpTLAYDate(0).SetFocus
        Exit Function
    Else
        If DaysBetween(rsEM("ED_DOH"), dlpTLAYDate(0)) < 1 Then
            MsgBox "Leave Date must be greater than Original Hire Date"
            rsEM.Close
            dlpTLAYDate(0).SetFocus
            Exit Function
        End If
    End If
rsEM.Close
End If

If Len(clpATTCode.Text) > 0 Then
    If clpATTCode.Caption = "Unassigned" Then
        MsgBox "Attendance code must be valid"
        clpATTCode.SetFocus
        Exit Function
    End If
    If Not clpATTCode.ListChecker Then
        'MsgBox "Attendance code must be valid"
        'clpATTCode.SetFocus
        Exit Function
    End If
End If
If Not glbLinamar Then
    If Len(clpCode(2).Text) < 1 Then
        MsgBox "Employment Status is a required field"
        clpCode(2).SetFocus
        Exit Function
    Else
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "Employment Status code must be valid"
            clpCode(2).SetFocus
            Exit Function
'        ElseIf Len(clpATTCode.Text) < 1 Then
'            MsgBox "Attendance Code is a required field"
'            clpATTCode.SetFocus
'            Exit Function
        Else
            EMPCode_Desc
            If chkLeave = 0 Then
                MsgBox "This is not code for Leave of Absence"
                clpCode(2).SetFocus
                Exit Function
            End If
            'Ticket #22965 Franks 12/17/2012
            '"   New Employment Status cannot equal current Employment Status.
            If OLDEMP = clpCode(2).Text Then
                MsgBox "New Employment Status cannot equal current Employment Status"
                clpCode(2).SetFocus
                Exit Function
            End If
        End If
    End If
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20648 Franks 09/23/2011
    If Len(comESalInc.Text) = 0 Then
        MsgBox "Eligible for Salary Increase is required."
        comESalInc.SetFocus
        Exit Function
    End If
    
End If

If glbWFC Then
'If glbWFC And frmWFCBenList.Visible Then 'Ticket #24061 Franks 07/16/20133
    If frmWFCBenList.Visible Then 'Ticket #24061 Franks 07/16/2013
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "' AND NOT (BM_ENDDATE IS NULL) "
        SQLQ = SQLQ & "AND BM_PCC = 1 " 'NEW - Company % only
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsTemp.EOF Then 'no End Date enter, then pop up this message
            Msg$ = "By entering this leave, do any benefits end sometime within the leave period?"
            Title$ = "Enter Leave Employee"
            DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
            Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
            
            If Response% = IDNO Then    ' Evaluate response
            '    Exit Function
            Else 'Yes
                'Ticket #24094 Franks 07/22/2013
                'o   If they answer YES, you cannot process the leave without them entering
                'an end date on some benefits. The current program doesn't stop the leave process when I enter YES.
                Exit Function
            End If
        End If
        rsTemp.Close
    End If
    'If Not frmWFCBenList.Visible Then 'Ticket #25248 Franks 03/24/2014
        If dlpLastDat2.Visible Then
            If clpCode(2).Text = "SALC" Then
                If glbUNION = "NONE" Or glbUNION = "EXEC" Then
                    If Len(dlpLastDat2.Text) = 0 Then
                        MsgBox lblTitle(7).Caption & " is mandatory for 'NONE' or 'EXEC' employees if New Status is 'SALC', default it as From Date"
                        dlpLastDat2.Text = dlpTLAYDate(0).Text
                        Exit Function
                    End If
                End If
            End If
        End If
    'End If
End If

chkTerms = True

End Function

Private Sub cll_EEFind(frmName As Form)

frmName.Enabled = True
frmName.lblEENum = ShowEmpnbr(glbLEE_ID)
frmName.lblEEID = glbLEE_ID
frmName.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
If glbLinamar Then
    frmName.Caption = "Temporary Lay Off - " & Left$(lblEEName, 10)
Else
    frmName.Caption = "Enter Leave - " & Left$(lblEEName, 10)
End If
End Sub

Private Sub chkAllDates_Click() 'Ticket #23920 Franks 07/03/2013
Dim SQLQ As String
Dim xID As Long
    If IsDate(dlpEndDate.Text) Then
        If Not Data2.Recordset.EOF Then
            xID = Data2.Recordset("BM_BENE_ID")
            If chkAllDates.Value Then 'checked
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = " & Date_SQL(dlpEndDate.Text) & " WHERE BM_WRKEMP = '" & glbUserID & "' "
                If clpCode(2).Text = "SALC" Then 'Ticket #30446 Franks 08/09/2017
                Else
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                End If
                gdbAdoIhr001.Execute SQLQ
                Data2.Refresh
                SQLQ = "BM_BENE_ID = " & xID
                Data2.Recordset.Find SQLQ
            Else 'unchecked
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = Null WHERE BM_WRKEMP = '" & glbUserID & "' "
                If clpCode(2).Text = "SALC" Then 'Ticket #30446 Franks 08/09/2017
                Else
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                End If
                SQLQ = SQLQ & "AND NOT (BM_BENE_ID = " & xID & ") "
                gdbAdoIhr001.Execute SQLQ
                Data2.Refresh
                SQLQ = "BM_BENE_ID = " & xID
                Data2.Recordset.Find SQLQ
            End If
        End If
    End If
End Sub

Private Sub clpCode_Change(Index As Integer)
If Index = 2 Then EMPCode_Desc
End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Sub cmdClose_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim SQLQ
Dim xEMP
Dim xPenStatus

If lblEEID = 0 Then
    MsgBox "        No Current Record" & Chr(10) & "Use 'FIND' to Select a Employee"
    Exit Sub
End If

If Not chkTerms() Then Exit Sub

Msg$ = Msg$ & Chr(10) & "Are you sure you want to "
If glbLinamar Then
    Msg$ = Msg$ & "Temporary Lay Off "
Else
    Msg$ = Msg$ & "Enter a Leave for "
End If
Msg$ = Msg$ & Chr(10) & "this employee?"
'Msg$ = Msg$ & Chr(10) & "Make sure no other info:HR Window "
'Msg$ = Msg$ & Chr(10) & "is open with this employee information showing"
If glbLinamar Then
    Title$ = "Temporary Lay Off Employee"
Else
    Title$ = "Enter Leave Employee"
End If
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If glbSamuel Then
    'Ticket #20885 Franks 11/18/2011
    xSamuleFlag = False
    Call CheckReptAuth
End If

'Ticket #24316 - WFC: If Status = 'SALC' then send Termination Email
If glbWFC And clpCode(2).Text = "SALC" Then
    If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'do not use it
        'If gsEMAIL_ONLEAVECHANGES Then
        If gsEMAIL_ONTERM Then 'Ticket #28970 Franks 07/27/2016  - send this terminated email is not based on gsEMAIL_ONLEAVECHANGES
            If Not WFC_TermEmailSending Then Exit Sub
        End If
        'End If
    End If
End If

MDIMain.panHelp(0).FloodType = 1

Screen.MousePointer = vbHourglass


If Not AUDITTERM() Then MsgBox "ERROR - AUDIT FILE"
Call updAttendance
Call updFollow
Call updStatus

If glbCompSerial = "S/N - 2351W" Then 'Ticket #17447 - Burlington Technologies only
'If Not glbLambton Then  'Ticket #15521
    Call updPositionHis 'Ticket #14623
End If

If glbWFC Then 'Ticket #16395 Pension System
    'If clpCode(2).Text = "TLAY" Then
        If WFCPensionEligible(lblEEID) Then
            'Ticket #19954 Franks 03/28/2011
            'Enter & extend a leave has no affect on the pension master
            'Call WFCPensionMasUpt(lblEEID, "Temporary Layoff", dlpTLAYDate(0).Text, , Year(CVDate(dlpTLAYDate(0).Text)))
            'Others -> STD
            'If clpCode(2).Text = "STD" Then
            '    If OLDEMP <> clpCode(2).Text Then
            '        If IsDate(dlpTLAYDate(0).Text) Then
            '            Call Upt_PENSIONDATE2(lblEEID, "UPDATE", dlpTLAYDate(0).Text)
            '        End If
            '    End If
            'End If
            
            'Ticket #21597 Franks 05/01/2012
            '"   Lookup the Table Master for the New Employment Status. If there is a Pension Status
            '(tb_usr1) entered and is different than the current Pension Status, update the Pension Master
            'with the Pension Status and Effective Date
            xPenStatus = getPenStatusFromHRTABL(clpCode(2).Text)
            
            'Ticket #21788 Franks 03/26/2012
            ''If clpCode(2).Text = "TLAY" Then
            'Ticket #21597 Franks 05/01/2012
            If clpCode(2).Text = "TLAY" Or Len(xPenStatus) > 0 Then
                'o   Update the Last Day (Status/Dates) to be the FROM DATE from Enter a Leave
                Call uptEmpDates(lblEEID, "ED_LDAY", dlpTLAYDate(0).Text)
                ''o   Update the Pension Status to "S"
                ''o   Update the Effective Date of Status to equal the FROM DATE from Enter a Leave
                ''o   Update the Benefit Rate to equal the dollar amount from the table
                'Call WFCPensionMasUpt(lblEEID, "Temporary Layoff", dlpTLAYDate(0).Text, , Year(CVDate(dlpTLAYDate(0).Text)))
                'Ticket #21597 Franks 05/01/2012 - add xPenStatus
                Call WFCPensionMasUpt(lblEEID, "Temporary Layoff", dlpTLAYDate(0).Text, xPenStatus, Year(CVDate(dlpTLAYDate(0).Text)))
            End If
            
            'Ticket #19954 Franks 03/28/2011
            'If an employee is on a leave, the Date of Disability (Pension Date 2) needs to be updated with the FROM date.
            If IsDate(dlpTLAYDate(0).Text) Then
                'Ticket #23361 Franks 03/11/2013
                If clpCode(2).Text = "STD" Or clpCode(2).Text = "LTD" Or clpCode(2).Text = "WCB" Then
                    Call Upt_PENSIONDATE2(lblEEID, "UPDATE", dlpTLAYDate(0).Text)
                End If
            End If
        End If
    'End If
    
    If xlocUSA = "Y" Then
        SQLQ = "UPDATE HREMP SET ED_PENSION = '2' "
        SQLQ = SQLQ & " WHERE ED_EMPNBR=" & lblEEID
        gdbAdoIhr001.Execute SQLQ
    End If
    
    'Ticket #19266 Franks 11/29/2010
    If dlpDOther2.Visible Then
        Call WFC_NGS_Trans(glbLEE_ID)
    End If
    
    'Ticket #23920 Franks 07/02/2013
    If frmWFCBenList.Visible Then 'US NGS employees
        xNewStatus = clpCode(2).Text 'Ticket #30446 Franks 08/09/2017 - xStatus
        Call WFC_NGSBenEndDateUpt(glbLEE_ID)
    End If
    'Ticket #25248 Franks 03/24/2014
    Call WFCNonNGSLastDayUpt(glbLEE_ID)
End If

If glbSamuel Then 'Ticket #20885 Franks 12/01/2011
    Call SAMUEL_Trans(glbLEE_ID)
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20648 Franks 09/23/2011
    Call EmployeeFlagUpd(lblEEID, 3, comESalInc.Text, dlpTLAYDate(0).Text, "", "")
End If

If Len(clpCode(2).Text) > 0 Then xEMP = clpCode(2).Text Else xEMP = "*"
'If Not EmpHisCalc(1, lblEEID, "", "", xEMP, "", "", "", "", Date) Then MsgBox "EMPHIS Error"
If Not EmpHisCalc(1, lblEEID, "", "", xEMP, "", "", "", "", Date, , , dlpTLAYDate(0).Text, dlpTLAYDate(1).Text) Then MsgBox "EMPHIS Error"

SQLQ = "UPDATE HREMP  SET "
SQLQ = SQLQ & fdFdate & "=" & Date_SQL(dlpTLAYDate(0).Text) & ", "
If Len(dlpTLAYDate(1).Text) > 0 Then 'Ticket #22009 Franks 05/10/2012
    SQLQ = SQLQ & fdTdate & "=" & Date_SQL(dlpTLAYDate(1).Text) & ", "
End If
SQLQ = SQLQ & " ED_EMP='" & clpCode(2).Text & "' "
SQLQ = SQLQ & " WHERE ED_EMPNBR=" & lblEEID

gdbAdoIhr001.Execute SQLQ

'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
If glbCompEntVacDaily Then
    Call Recompute_DailyAccrualFile(glbLEE_ID, dlpTLAYDate(0).Text)
End If

'Ticket #18368
If glbCompSerial = "S/N - 2259W" Or glbCompSerial = "S/N - 2241W" Then
    Call Employee_Master_Integration(glbLEE_ID)
End If
'Ticket #19071 GP Frontenac
If glbCompSerial = "S/N - 2410W" Then
    Call Employee_Master_Integration(glbLEE_ID, , , , "LOA")
End If
If glbWFC Then 'Ticket #25116 Franks 02/25/2014
    If glbAdv Then
        Call Employee_Master_Integration(glbLEE_ID)
    End If
End If

'Send Email
If gsEMAIL_ONLEAVECHANGES Then
    MailBody = ""
    If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
        MailBody = GetEmailBodyForSamuel(glbLEE_ID)
        MailBody = MailBody & "is on Leave of Absence." & vbCrLf & vbCrLf
        MailBody = MailBody & "Pay Status: " & GetTABLDesc("EDEM", clpCode(2)) & vbCrLf
        MailBody = MailBody & "Pay Status Change Effective (From): " & dlpTLAYDate(0) & vbCrLf
        If Len(MailBody) > 0 Then
           Screen.MousePointer = DEFAULT
           Call EmailSendingForSamuel
        End If
    Else
        MailBody = "The employee is on Leave of Absence." & vbCrLf & vbCrLf
        MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
        MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
        MailBody = MailBody & "Reason: " & GetTABLDesc("EDEM", clpCode(2)) & vbCrLf
        If Len(dlpTLAYDate(1).Text) = 0 Then 'Ticket #22009 Franks 05/10/2012
            MailBody = MailBody & "Effective From: " & dlpTLAYDate(0) & vbCrLf
        Else
            MailBody = MailBody & "Effective From: " & dlpTLAYDate(0) & " To: " & dlpTLAYDate(1) & vbCrLf
        End If
        If Len(MailBody) > 0 Then
           Screen.MousePointer = DEFAULT
           Call imgEmail_Click
        End If
    End If
End If

If glbSamuel Then 'Ticket #25883 Franks 08/29/2014
    Call CheckReptDispRpt
End If

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).FloodPercent = 100
oFDate = dlpTLAYDate(0).Text
OTDate = dlpTLAYDate(1).Text
'cmdOK.Enabled = False
'cmdClose.SetFocus
MDIMain.panHelp(0).FloodType = 0

If glbSamuel And xSamuleFlag Then 'Ticket #25883 Franks 08/29/2014
    'dont leave this screen, otherwise the report will be closed automatically for samuel
Else
    Unload Me 'Ticket #24061 Franks 07/16/2013
End If

End Sub

Private Sub WFCPensionUpdate(xEmpNo, xFDate)
Dim rsPen As New ADODB.Recordset
Dim SecCode, UnionCode, SalHrl
Dim SQLQ As String
Dim PenType As String
    SecCode = GetEmpData(xEmpNo, "ED_SECTION")
    UnionCode = GetEmpData(xEmpNo, "ED_ORG")
    SalHrl = GetSalHourly(SecCode, UnionCode)
    PenType = GetPensionType(SecCode, UnionCode)
    'Update the Pension Master with the DB Status equal to "S"
    'and the DB Status Effective Date = From Date of the Employment Status.
    If SalHrl = "Hourly" And Len(PenType) > 0 Then
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_EMPNBR = " & xEmpNo & " "
        If Len(SecCode) > 0 Then
            SQLQ = SQLQ & "AND PE_SECTION = '" & SecCode & "' "
        End If
        SQLQ = SQLQ & "AND PE_PENSIONTYPE='" & PenType & "' "
        SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC, PE_PENSIONTYPE "
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsPen.EOF Then
            rsPen("PE_DB_STATUS") = "S"
            If IsDate(xFDate) Then
                rsPen("PE_DB_STATUS_DATE") = xFDate
            End If
            rsPen.Update
        End If
        rsPen.Close
    End If
End Sub
'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(Me.ActiveControl) '19Aug99 js
'End Sub

Private Sub EMPCode_Desc()
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
On Error GoTo EMPCode_Desc_Err
chkLeave.Value = 0

If Len(clpCode(2).Text) > 0 Then
    SQLQ = "SELECT TB_USR3 FROM HRTABL WHERE TB_NAME='EDEM' AND TB_KEY = '" & clpCode(2).Text & "'"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsTA.EOF Then
        chkLeave.Value = IIf(rsTA("TB_USR3"), 1, 0)
    End If
End If

Exit Sub
EMPCode_Desc_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EMP Code Snap", "TABL", "SELECT")
Call RollBack '29July99 js

End Sub

Private Sub clpCode_LostFocus(Index As Integer)
If glbWFC Then 'Ticket #26308 Franks 11/27/2014
    If Index = 2 Then
        Call WFCLastDaySetup
        
        If clpCode(2).Text = "SALC" Then 'Ticket #30446 Franks 08/09/2017
            chkAllDates.Value = 1
            dlpEndDate.Text = dlpTLAYDate(0).Text
            Call chkAllDates_Click
        End If
    End If
    
End If
End Sub

Private Sub WFCLastDaySetup() 'Ticket #26308 Franks 11/27/2014
    If dlpLastDat2.Visible Then
        If clpCode(2).Text = "SALC" Then
            If IsDate(dlpTLAYDate(0).Text) Then
                dlpLastDat2.Text = dlpTLAYDate(0).Text
            End If
        End If
    End If
End Sub

Private Sub clpJob_LostFocus()
Dim xNSalary, xNOrg, xNDHRS, xNWHRS
Dim TE As New ADODB.Recordset
Dim SQLQ

xNSalary = 0
xNOrg = ""
xNDHRS = 0
xNWHRS = 0

    SQLQ = "SELECT SH_SALCD,SH_SALARY FROM HR_SALARY_HISTORY WHERE SH_EMPNBR=" & glbLEE_ID & " AND SH_JOB='" & clpJob & "' AND SH_CURRENT<>0 "
    TE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not TE.EOF Then
        If TE("SH_SALCD") = "A" Then
            comPayPer.ListIndex = 0
        ElseIf TE("SH_SALCD") = "H" Then
            comPayPer.ListIndex = 1
        ElseIf TE("SH_SALCD") = "D" And glbCompSerial = "S/N - 2282W" Then
            comPayPer.ListIndex = 3
        Else
            comPayPer.ListIndex = 2
        End If
        xNSalary = TE("SH_SALARY")
    End If
    TE.Close
    SQLQ = "SELECT JH_ORG,JH_DHRS,JH_WHRS,JH_PAYROLL_ID,JH_GLNO FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & glbLEE_ID & " AND JH_JOB='" & clpJob & "' AND JH_CURRENT<>0 "
    TE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not TE.EOF Then
        xNOrg = IIf(IsNull(TE("JH_ORG")), clpCode(0), TE("JH_ORG"))
        xNDHRS = TE("JH_DHRS")
        xNWHRS = TE("JH_WHRS")
        medSalary = xNSalary
        clpCode(0) = xNOrg
        txtDHRS = xNDHRS
        txtWHRS = xNWHRS
        clpGLNo = TE("JH_GLNO") & ""
        txtPayrollID = TE("JH_PAYROLL_ID") & ""
    End If
    TE.Close

End Sub

Private Sub Job_Desc()
Dim SQLQ As String
On Error GoTo Pos_Err
Dim dynaJobs As New ADODB.Recordset
 clpJob = ""
 clpJob.ShowDescription = False
If Len(clpJob.Text) > 0 Then
     clpJob.Caption = "Unassigned"
     clpJob.ShowDescription = True
    dynaJobs.Open "HRJOB", gdbAdoIhr001, adOpenDynamic
    If dynaJobs.EOF And dynaJobs.BOF Then Exit Sub
    SQLQ = "JB_CODE = '" & clpJob.Text & "'"
    dynaJobs.Find SQLQ
    If Not dynaJobs.EOF Then clpJob.Caption = dynaJobs("JB_DESCR")
End If

Exit Sub

Pos_Err:
If Err = 94 Then
    Err = 0
    Resume Next
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "JOBS", "SELECT")
Call RollBack '28July99 js

End Sub

Private Sub cmdImport_Click()
    'Make sure the LOA Date Range and the new Employment Status is entered before selecting the document
    If Not chkTerms() Then Exit Sub
    
    glbDocNewRecord = False
    glbDocName = "LOA"
    'glbDocKey = rsDATA("SC_ID")
    If fglbNew Then
        glbDocKey = 0
    'Else
    '    glbDocKey = RSDATA("SC_ID") 'Ticket #16018
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmETLAY")
End Sub

Private Sub cmdPostion_Click()
Dim oJob As String, OJobD As String

oJob = clpJob.Text
OJobD = clpJob.Caption

Load frmJOBS
frmJOBS.Show 1

'If Len(glbJob) < 1 Then
If Len(glbPos) < 1 Then
    clpJob.Text = oJob
    clpJob.Caption = OJobD
Else
    clpJob.Text = glbPos
    clpJob.Caption = glbPosDesc
    
    'clpJob.Text = glbJob
    'clpJob.Caption = glbJobDesc
End If
End Sub

Private Sub comPayPer_LostFocus()
If comPayPer.ListIndex = 0 Then lblSalCode.Caption = "A"
If comPayPer.ListIndex = 1 Then lblSalCode.Caption = "H"
If comPayPer.ListIndex = 2 Then lblSalCode.Caption = "M"
If comPayPer.ListIndex = 3 Then lblSalCode.Caption = "D"
End Sub

Private Sub dlpENDDATE_LostFocus()
If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
    If Data2.Recordset("BM_PCE") = 1 Then 'employee %
        If IsNull(Data2.Recordset("BM_ENDDATE")) Then
            If IsDate(dlpEndDate.Text) Then
                MsgBox "Cannot enter END DATE on 100% paid employee benefits."
                dlpEndDate.Text = ""
                Exit Sub
            End If
        End If
    End If
End If
Call WFCUpdate_Value
End Sub

Private Sub dlpTLAYDate_LostFocus(Index As Integer)
If glbLinamar Then
    If Index = 0 Then
        If IsDate(dlpTLAYDate(0)) And Not IsDate(dlpTLAYDate(1)) Then
            dlpTLAYDate(1) = DateAdd("ww", 35, dlpTLAYDate(0))
        End If
    End If
End If
If IsDate(dlpTLAYDate(1)) And IsDate(dlpTLAYDate(0)) Then
    lblWeeks = Format((CVDate(dlpTLAYDate(1)) - CVDate(dlpTLAYDate(0))) / 7, "###.0") & " Weeks"
Else
    lblWeeks = ""
End If
If glbWFC Then 'Ticket #26308 Franks 11/27/2014
    Call WFCLastDaySetup
End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMETLAY"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMETLAY"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim SQLQ As String

glbOnTop = "FRMETLAY"
glbchkSum = False  'Jaddy 11/9/99

If glbLinamar Then
    fdFdate = "ED_USRDAT1"
    fdTdate = "ED_UNION"
    Me.Caption = "Temporary Lay Off"
    lblTitle(1) = "Temporary Lay Off Reason"

    'Release 8.1
    lblComment.Visible = True
    lblComment.Top = 4440
    memComments.Visible = True
    memComments.Top = 4740
Else
    fdFdate = "ED_SFDATE"
    fdTdate = "ED_STDATE"
    Me.Caption = "Enter Leave"
    lblTitle(1) = "Enter Leave Reason"
    lblTitle(1).Visible = False
    clpCode(1).Visible = False
    clpCode(1).ShowDescription = False
    
    'Release 8.1
    lblComment.Visible = True
    lblComment.Top = 3840
    memComments.Visible = True
    memComments.Top = 4140
    
    'George on Jan 26,2006 #10266
    If gsAttachment_DB Then
        glbJob = ""
        glbSDate = "01/01/1900"
        lblImport.Visible = True 'False
        imgSec.Visible = False
        imgNoSec.Visible = True 'False
        cmdImport.Visible = True 'False
    End If
    'George on Jan 26,2006 #10266
End If

'ticket #16952
If glbCompSerial = "S/N - 2390W" Then
    If Not gSec_Inq_Salary Then
        frmMulti.Visible = False
    End If
End If

If glbNoNONE Or glbNoEXEC Then 'Ticket #28512 Franks 04/20/2016
    If glbUNION = "NONE" Or glbUNION = "EXEC" Then
        frmMulti.Visible = False
    End If
End If

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

If glbAdv Then 'Ticket #14739
    If glbWFC Then
    'WFC not use this function, so hide this checkbox
    Else
        chkATPaidHours.Visible = True
    End If
End If

Screen.MousePointer = DEFAULT

If glbLEE_ID = 0 And (Not glbtermopen) Then frmEEFIND.Show 1
If glbLEE_ID > 0 Then
    Me.Show
    Call cll_EEFind(Me)
Else
    Unload Me
    Exit Sub
End If
If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
    Me.Left = 0
End If
Me.WindowState = vbMaximized
lblWeeks = ""

If glbMulti Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2363W" Then
    comPayPer.Clear
    comPayPer.AddItem "Annum"
    comPayPer.AddItem "Hour"
    comPayPer.AddItem "Monthly"
    'If glbCompSerial = "S/N - 2282W" Then
    '    comPayPer.AddItem "Daily "
    'End If
    frmMulti.Visible = True
Else
    comPayPer.Clear
    comPayPer.AddItem "Annum"
    comPayPer.AddItem "Hour"
    comPayPer.AddItem "Monthly"
    If glbWFC Then
        comPayPer.AddItem "Daily "
    End If
End If

Call EERetrieve

Call CR_JobHis_Snap 'Ticket #13165

'Get current salary FOR Casey House
Dim rsSalT As New ADODB.Recordset
SQLQ = "SELECT SH_EMPNBR,SH_CURRENT,SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & glbLEE_ID
rsSalT.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Not rsSalT.EOF Then
    medSalary.Text = rsSalT("SH_SALARY")
    If Not IsNull(rsSalT("SH_SALCD")) Then 'Ticket #15081
        lblSalCode.Caption = rsSalT("SH_SALCD")
    End If
Else
    medSalary.Text = 0
    comPayPer = ""
    lblSalCode.Caption = "" 'Ticket #13312
End If
rsSalT.Close

SQLQ = "SELECT ED_EMPNBR, ED_DEPTNO,ED_GLNO,ED_ADMINBY,ED_ORG,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
clpCode(0) = "": clpGLNo = ""
rsSalT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsSalT.EOF Then
    If Not IsNull(rsSalT("ED_GLNO")) Then
        clpGLNo.Text = rsSalT("ED_GLNO")
    End If
    If Not IsNull(rsSalT("ED_ORG")) Then
        clpCode(0).Text = rsSalT("ED_ORG")
    End If
    If Not IsNull(rsSalT("ED_PAYROLL_ID")) Then
        txtPayrollID = rsSalT("ED_PAYROLL_ID")
    End If
End If
rsSalT.Close

SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & glbLEE_ID
clpJob = "": txtDHRS = 0: txtWHRS = 0
rsSalT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsSalT.EOF Then
    clpJob.Text = rsSalT("JH_JOB")
    If Not IsNull(rsSalT("JH_DHRS")) Then txtDHRS = rsSalT("JH_DHRS")
    If Not IsNull(rsSalT("JH_WHRS")) Then txtWHRS = rsSalT("JH_WHRS")
End If
rsSalT.Close

'***************
If glbCompSerial = "S/N - 2192W" Then  ' county essex
    'do not show job
    cmdPostion.Visible = False
    lblTitle(8).Visible = False
    clpJob.Visible = False
    ' union
    lblTitle(9).Visible = False
    clpCode(0).Visible = False
    
    lblTitle(10).Top = 270
    medSalary.Top = 270
    
    lblTitle(11).Top = 570
    comPayPer.Top = 570
    
    lblHrsDay.Visible = False
    txtDHRS.Visible = False
    lblHrsWeek.Visible = False
    txtWHRS.Visible = False
    lblTitle(20).Visible = False
    clpGLNo.Visible = False
    lblPayID.Visible = False
    txtPayrollID.Visible = False
   
    frmMulti.Height = 1000
ElseIf glbCompSerial = "S/N - 2366W" Then 'FYC Muskoka
    'do not show job
    cmdPostion.Visible = False
    lblTitle(8).Visible = False
    clpJob.Visible = False
    ' union
    lblTitle(9).Visible = False
    clpCode(0).Visible = False
    ' Hours/day
    lblHrsDay.Visible = False
    txtDHRS.Visible = False
    lblHrsWeek.Visible = False
    txtWHRS.Visible = False
    'GLNo
    lblTitle(20).Visible = False
    clpGLNo.Visible = False
    lblPayID.Visible = False
    txtPayrollID.Visible = False
End If
'***************

If glbWFC Then 'Ticket #19266 Franks 11/29/10
    Call WFCOther2Screen(glbLEE_ID)
    lblTitle(2).FontBold = False 'Ticket #22009 Franks 05/10/2012
    Call WFCBenListScreen(glbLEE_ID) 'Ticket #23920 Franks 07/02/2013
    Call WFCLastDayForNonNGSEmployee(glbLEE_ID) 'Ticket #25248 Franks 03/24/2014
    
    'Release 8.1
    If frmWFCBenList.Visible = True Then
        lblComment.Visible = True
        lblComment.Top = 6840 + 200
        'lblComment.Left = 480
        memComments.Visible = True
        memComments.Top = 7140 + 200
        'memComments.Left = 480
    End If
    
    clpJob.TextBoxWidth = 1315 'Ticket #28118 Franks 02/02/2016
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20648 Franks 09/23/2011
    Call SamuelScreenSetup
End If

Screen.MousePointer = HOURGLASS

If glbLinamar Then
    MDIMain.panHelp(0).Caption = "Proceed with Temporary Lay Off"
Else
    MDIMain.panHelp(0).Caption = "Proceed with Enter Leave"
End If
If Not gSec_Upd_EnterLeave Then
'    cmdOK.Enabled = False
    dlpTLAYDate(0).Enabled = False
    dlpTLAYDate(1).Enabled = False
    clpCode(1).Enabled = False
    clpCode(2).Enabled = False
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call INI_Controls(Me)
If glbLinamar Then
    clpCode(1).Text = "TLAY"
    clpCode(1).Caption = "Temporary Leaves"
    clpCode(2).Text = "TEMP"
    clpATTCode.Text = "TLAY"
End If

clpJob.seleEMPCode = fglbJobList 'Ticket #13165

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
'Ticket #23920 Franks 07/02/2013 - don't need to check change, it caused an error on frmWFCBenList
'Keepfocus = Not isUpdated(Me)
'Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized And Me.WindowState <> vbMaximized Then
        Me.Left = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmETLAY = Nothing
End Sub

Private Function ReadJob()
Dim rsTA As New ADODB.Recordset
Dim IJob
ReadJob = ""

rsTA.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & lblEEID, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Function
ReadJob = rsTA("JH_JOB")
rsTA.Close

End Function

'Private Sub txtTLAYDate_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtTLAYDate_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtTLAYDate_GotFocus(Index As Integer)
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

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

Private Sub updFollow()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim xReasonDesc As String

On Error GoTo CrFollow_Err

'Ticket #22009 Franks 05/10/2012
'No To Date then no Follow up Effetive Date, so don't create follow up, this is for wfc only now.
If Len(dlpTLAYDate(1).Text) = 0 Then
    Exit Sub
End If

SQLQ = "SELECT * FROM HR_FOLLOW_UP "
If IsDate(OTDate) Then
    SQLQ = SQLQ & " WHERE EF_COMPLETED=0 AND EF_EMPNBR=" & lblEEID
    If glbLinamar Then
        SQLQ = SQLQ & " AND EF_FREAS='TLAY' "
    Else
        SQLQ = SQLQ & " AND EF_FREAS='LOA' "
    End If
    SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(OTDate)
End If

rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not IsDate(OTDate) Or rsTB.EOF Then
    rsTB.AddNew
    Msg = "A Follow Up Record was created!"
Else
    Msg = "A Follow Up Record was updated!"
End If
rsTB("EF_COMPNO") = "001"
rsTB("EF_EMPNBR") = lblEEID
rsTB("EF_FDATE") = CVDate(dlpTLAYDate(1).Text)
rsTB("EF_FREAS_TABL") = "FURE"
'Ticket #24257 - Do not update Admin By for them only
If glbCompSerial <> "S/N - 2262W" Then
    rsTB("EF_ADMINBY_TABL") = "EDAB"
    rsTB("EF_ADMINBY") = GetEmpData(lblEEID, "ED_ADMINBY", Null)
End If
If glbLinamar Then
    rsTB("EF_FREAS") = "TLAY"
Else
    rsTB("EF_FREAS") = "LOA"
    Dim rsTT As New ADODB.Recordset
    rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='LOA'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsTT.EOF Then
        rsTT.AddNew
        rsTT("TB_COMPNO") = "001"
        rsTT("TB_NAME") = "FURE"
        rsTT("TB_KEY") = "LOA"
        rsTT("TB_DESC") = "Leave of Absence Review"
        rsTT("TB_LUSER") = glbUserID
        rsTT("TB_LDATE") = Date
        rsTT("TB_LTIME") = Time$
        rsTT.Update
    End If
    rsTT.Close
    
    'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
    'follow up record
    Call Grant_FollowUpCode_Security(glbUserID, "LOA", "Leave of Absence Review")
    
End If

'Get the Employment Status for LOA description
xReasonDesc = GetTABLDesc("EDEM", clpCode(2))

If glbLinamar Then
    rsTB("EF_COMMENTS") = lblEEName & " was " & IIf(glbLinamar, "temporarily laid off on ", "on leave from ") & Format(dlpTLAYDate(0).Text, "mmmm dd, yyyy")
Else
    rsTB("EF_COMMENTS") = xReasonDesc & ": " & lblEEName & " was " & IIf(glbLinamar, "temporarily laid off on ", "on leave from ") & Format(dlpTLAYDate(0).Text, "mmmm dd, yyyy")
End If
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update


fglbFollowID = rsTB("EF_FOLLOWUP_ID")
rsTB.Close
MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Private Sub updStatus()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xID

On Error GoTo CrFollow_Err

rsTA.Open "SELECT ED_EMP FROM HREMP WHERE ED_EMPNBR=" & lblEEID, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Sub

SQLQ = "SELECT * FROM HRSTATUS "
If IsDate(oFDate) Or IsDate(OTDate) Then
    SQLQ = SQLQ & " WHERE SC_REASON IN ('TLAY', 'LOA') AND SC_EMPNBR=" & lblEEID
    If IsDate(oFDate) Then SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(oFDate)
    If IsDate(OTDate) Then SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(OTDate)
    SQLQ = SQLQ & " AND SC_TYPE='HR'"
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsTB.EOF And (IsDate(oFDate)) Then
    rsTB("SC_TYPE") = Null
    rsTB.Update
End If

rsTB.AddNew
rsTB("SC_COMPNO") = "001"
rsTB("SC_EMPNBR") = lblEEID
rsTB("SC_FDATE") = dlpTLAYDate(0).Text
If Len(dlpTLAYDate(1).Text) > 0 Then 'Ticket #22009 Franks 05/10/2012
rsTB("SC_TDATE") = dlpTLAYDate(1).Text
End If
rsTB("SC_EMP_TABL") = "EDEM"
rsTB("SC_OLDEMP") = rsTA!ED_EMP
rsTB("SC_NEWEMP") = clpCode(2).Text
rsTB("SC_REASON_TABL") = "SCRE"
If glbLinamar Then
    rsTB("SC_REASON") = clpCode(1).Text
Else
    rsTB("SC_REASON") = "LOA"
End If
rsTB("SC_FOLLOWID") = fglbFollowID

If Len(clpATTCode.Text) > 0 Then rsTB("SC_ATTREASON") = clpATTCode.Text

rsTB("SC_JOB") = ReadJob
rsTB("SC_TYPE") = "HR"
rsTB("SC_LDATE") = Date
rsTB("SC_LTIME") = Time$
rsTB("SC_LUSER") = glbUserID

'Release 8.1
rsTB("SC_COMMENT") = memComments.Text

rsTB.Update

'Release 8.1
xID = rsTB("SC_ID")

rsTB.Close

'Release 8.1
If gsAttachment_DB Then
    'If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            'If glbtermopen Then
            '    Call AttachmentAdd(glbTERM_ID, glbDocImpFile, glbDocType, glbDocDesc)
            'Else
                'Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
                
                gdbAdoIhr001_DOC.BeginTrans
                gdbAdoIhr001_DOC.Execute "Update HRDOC_HRSTATUS set SC_DOCKEY = " & glbDocKey & ", SC_STYPE='" & clpCode(2).Text & "', SC_FDATE=" & Date_SQL(dlpTLAYDate(0).Text) & " WHERE SC_TYPE='" & UCase(glbDocName) & "' AND SC_EMPNBR = " & glbLEE_ID & " AND SC_DOCKEY = 0 AND SC_DOCTYPE = '" & glbDocType & "' AND SC_USRDESC = '" & glbDocDesc & "'"
                gdbAdoIhr001_DOC.CommitTrans
                
                SQLQ = "UPDATE HRSTATUS SET SC_DOCKEY = SC_ID WHERE SC_ID=" & glbDocKey
                gdbAdoIhr001.Execute SQLQ

            'End If
        End If
    'End If
    glbDocImpFile = ""
End If

Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Function EERetrieve()
Dim SQLQ As String
Dim rsEmp As New ADODB.Recordset
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

OLDEMP = "" 'Ticket #22965
SQLQ = "Select ED_EMPNBR,ED_USRDAT1,ED_UNION,ED_SFDATE,ED_STDATE,ED_EMP from HREMP"
SQLQ = SQLQ & " where ED_EMPNBR=" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'dlpTLAYDate(0).text = "": dlpTLAYDate(1).text = ""
If Not rsEmp.EOF Then
    If glbLinamar Then
        If IsDate(rsEmp(fdFdate)) Then
            dlpTLAYDate(0).Text = rsEmp(fdFdate)
            If glbLinamar Then lblNotice.Visible = True
            If IsDate(rsEmp(fdTdate)) Then dlpTLAYDate(1).Text = rsEmp(fdTdate)
        End If
    End If
    If Not IsNull(rsEmp("ED_EMP")) Then
        OLDEMP = rsEmp("ED_EMP") 'Ticket #22965
    End If
End If
oFDate = dlpTLAYDate(0).Text
OTDate = dlpTLAYDate(1).Text
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HRDOLENT", "SELECT")
Resume Next

End Function

Sub Display_Value()
    Call cll_EEFind(Me)
    Call EERetrieve
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
RelateMode = RelateTermEmp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_EnterLeave
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
'ElseIf rsEMP.EOF Then
'    UpdateState = NoRecord
'    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
'    frmEBENEFITS.Caption = "Benefits / Beneficiaries - " & Left$(glbLEE_SName, 5)
    frmETLAY.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Function updPositionHis()
Dim SQLQ As String
    If IsDate(dlpTLAYDate(1).Text) Then
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_ENDDATE = " & Date_SQL(dlpTLAYDate(1).Text) & " "
        SQLQ = SQLQ & "WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & glbLEE_ID
        gdbAdoIhr001.Execute SQLQ
    End If
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_ENDREAS = '" & clpCode(2).Text & "' "
    SQLQ = SQLQ & "WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & glbLEE_ID
    gdbAdoIhr001.Execute SQLQ
        
End Function

Private Function updAttendance()
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset, rsDup As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsATT As New ADODB.Recordset
Dim xDays
Dim X, xDATE, xDup
Dim WSQLQ, ESQLQ, Result
Dim TSQLQ
Dim Msg$
Dim AskWeekend, SkipWeekend
Dim xWeekDay
Dim xAskDup
Dim xAddDup
Dim xKey

Dim xHours, xSHIFT, xSuper, xIncID, xSEN, xEMELEA, xINDICATOR
updAttendance = False
On Error GoTo updAttendance_Err

If Len(clpATTCode.Text) = 0 Then Exit Function
Screen.MousePointer = HOURGLASS

xHours = 0
xSHIFT = Null
xSuper = Null
rsJOB.Open "SELECT JH_DHRS,JH_REPTAU,JH_SHIFT FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID & " AND JH_JOB='" & clpJob.Text & "'", gdbAdoIhr001, adOpenForwardOnly
If Not rsJOB.EOF Then
    If IsNumeric(rsJOB("JH_DHRS")) Then xHours = rsJOB("JH_DHRS") Else xHours = 0
    xSuper = rsJOB("JH_REPTAU")
    xSHIFT = rsJOB("JH_SHIFT")
End If
rsJOB.Close
rsTB.Open "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY='" & clpATTCode.Text & "'", gdbAdoIhr001, adOpenForwardOnly
xIncID = 0
xSEN = 0
xEMELEA = 0
xINDICATOR = 0
If Not rsTB.EOF Then
    xSEN = rsTB("TB_SEN")
    xEMELEA = rsTB("TB_USR3")
    xINDICATOR = rsTB("TB_INDICATOR")
End If
rsTB.Close

If UCase(clpATTCode.Text) = "OT15" Then xHours = xHours * 1.5
If UCase(clpATTCode.Text) = "OT20" Then xHours = xHours * 2

'City of Timmins - Ticket #16168
If glbCompSerial = "S/N - 2375W" Then
    If UCase(clpATTCode.Text) = "OT05" Then xHours = xHours * 0.5
    If UCase(clpATTCode.Text) = "OT25" Then xHours = xHours * 2.5
End If

If Len(dlpTLAYDate(1).Text) = 0 Then
    xDays = 0
Else
    xDays = DateDiff("d", dlpTLAYDate(0).Text, dlpTLAYDate(1).Text)
End If

xDATE = dlpTLAYDate(0).Text
AskWeekend = True
xAskDup = True
xAddDup = True

For X = 0 To xDays
   xWeekDay = Weekday(xDATE)
   If xWeekDay = 7 Or xWeekDay = 1 Then
        If AskWeekend Then
            Msg$ = "Do you want exclude Saturday/Sunday for Attendance Records?"
            AskWeekend = False
            SkipWeekend = False
            If MsgBox(Msg$, 36) = 6 Then
                SkipWeekend = True
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
            AskWeekend = False
        Else
            If SkipWeekend Then
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
        End If
    End If
    If Len(dlpTLAYDate(1).Text) > 0 Then
        If CVDate(xDATE) > CVDate(dlpTLAYDate(1).Text) Then Exit For
    Else
        If CVDate(xDATE) > CVDate(dlpTLAYDate(0).Text) Then Exit For
    End If
       
    TSQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE "
    TSQLQ = TSQLQ & " WHERE AD_REASON = '" & clpATTCode.Text & "' "
    TSQLQ = TSQLQ & " AND AD_DOA = " & Date_SQL(xDATE)
    TSQLQ = TSQLQ & " AND AD_EMPNBR =" & glbLEE_ID
    rsDup.Open TSQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsDup.EOF Then
        If xAskDup Then
            Msg$ = "Reason: " & clpATTCode & Chr(10) & " Date: " & xDATE & Chr(10) & Chr(10)
            Msg$ = Msg$ & rsDup.RecordCount & " duplicates found in Attendance Master. " & Chr(10) & Chr(10)
            Msg$ = Msg$ & "Click Yes to post all Attendance records including duplicates." & Chr(10)
            Msg$ = Msg$ & "Click No to post all non-duplicate Attendance records." & Chr(10)
            Msg$ = Msg$ & "Click Cancel to cancel posting of all Attendance records." & Chr(10)
            Result = MsgBox(Msg$, vbYesNoCancel, "Duplicates Found")
            If Result = vbYes Then
                xDup = True
                xAddDup = True
            ElseIf Result = vbNo Then
                xDup = True 'False
                xAddDup = False
            ElseIf Result = vbCancel Then
                Exit For
            End If
            xAskDup = False
        End If
    Else
        xDup = False 'True
    End If
    rsDup.Close
    
    
    If Not xDup Or (xDup And xAddDup) Then
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=0"
        rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        rsATT.AddNew
        rsATT("AD_EMPNBR") = glbLEE_ID
        rsATT("AD_COMPNO") = "001"
        rsATT("AD_DOA") = xDATE
        rsATT("AD_REASON") = clpATTCode.Text
        rsATT("AD_HRS") = xHours
        rsATT("AD_SHIFT") = xSHIFT
        rsATT("AD_SUPER") = xSuper
        rsATT("AD_INCID") = xIncID
        rsATT("AD_SEN") = xSEN
        rsATT("AD_EMELEA") = xEMELEA
        rsATT("AD_INDICATOR") = xINDICATOR
        rsATT("AD_JOB") = clpJob.Text
        rsATT("AD_ORG") = clpCode(0).Text
        If IsNumeric(medSalary.Text) Then
            rsATT("AD_SALARY") = medSalary.Text
        End If
        rsATT("AD_DHRS") = txtDHRS
        rsATT("AD_WHRS") = txtWHRS
        rsATT("AD_GLNO") = clpGLNo.Text
        rsATT("AD_PAYROLL_ID") = txtPayrollID
        rsATT("AD_SALCD") = lblSalCode.Caption
        rsATT("AD_LDATE") = Date
        rsATT("AD_LTIME") = Time$
        rsATT("AD_LUSER") = glbUserID
        rsATT.Update

        If glbAdv Then 'Ticket #14739
            xKey = rsATT("AD_EMPNBR")
            xKey = xKey & "|" & Format(rsATT("AD_DOA"), "dd-mmm-yyyy")
            xKey = xKey & "|" & rsATT("AD_REASON")
            If chkATPaidHours.Value Then
                Call Attendance_Master_Integration(xKey, rsATT("AD_ATT_ID"))
            Else
                Call Attendance_Master_Integration(xKey, rsATT("AD_ATT_ID"), , "YES")
            End If
        End If
        
        rsATT.Close
    End If
    
    xDATE = DateAdd("d", 1, xDATE)
Next
Call EntReCalc("ED_EMPNBR=" & glbLEE_ID)
Call EntReCalcHr


updAttendance = True

Exit Function

updAttendance_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "updAttendance", "Attendance", "Insert")
updAttendance = False
Resume Next

End Function

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

'fUPMode = TF    ' update mode

clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
dlpTLAYDate(0).Enabled = TF
dlpTLAYDate(1).Enabled = TF
clpATTCode.Enabled = TF

End Sub

Private Sub lblSalCode_Change()
'If glbMulti Or glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2366W" Or glbCompSerial = "S/N - 2363W" Then
    If Len(lblSalCode) > 0 Then
        If lblSalCode = "A" Then
            comPayPer.ListIndex = 0
        ElseIf lblSalCode = "H" Then
            comPayPer.ListIndex = 1
        ElseIf lblSalCode = "M" Then
            comPayPer.ListIndex = 2
        ElseIf lblSalCode = "D" Then
            comPayPer.ListIndex = 3
        End If
    Else
        comPayPer = ""
    End If
'End If
End Sub

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    If gsEMAIL_ONLEAVECHANGES Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            
        'Ticket #18235 - Email on Leave Changes
        xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
        End If
        'Ticket #18235 - End
                    
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            'frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice"
            'Ticket #18578
            'frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice - " & lblEEName
            'Ticket #18755
            xBranch = GetEmpData(glbLEE_ID, "ED_SECTION", "")
            If Len(xBranch) > 0 Then
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR Leave Changes Notice - " & xBranch & lblEEName
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
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    'Resume Next
    Exit Sub

End Sub

Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
    If gsEMAIL_ONLEAVECHANGES Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            
        'Ticket #18235 - Email on Leave Changes
        If glbCompSerial = "S/N - 2382W" Then  'Samuel
            xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            End If
        Else
            'Ticket #20317 - More Emails for everyone
            xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            End If
        End If
        'Ticket #18235 - End
            
        'If Len(xEmail) > 0 Then    'Hemu - (Ticket #11562) - Jerry asked to remove the check for email address presence.
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352
            Else
                frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            End If
            frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice"
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            'End If
        '    MsgBox "There is no email address for the 'Email Notification on Leave Changes' on Company Preference screen. "
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

Private Sub WFCOther2Screen(xEmpNo)
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
    
    'chkATPaidHours.Top = 2640 + 400 'Ticket #24822 Franks 12/18/2013
    chkATPaidHours.Top = 2640 + 400 + 400 'Ticket #27045 Franks 05/08/2015
    
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
    End If
    rsEmpee.Close
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    
    xNGSStart = ""
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close
    'No NGS Effective Date, skip
    If Len(xNGSStart) = 0 Then Exit Sub
    lbOtherDate2.Caption = lStr("Other Date 2")
    lbOtherDate2.Top = Label1.Top + 420
    dlpDOther2.Top = Label1.Top + 420
    dlpDOther2.Left = clpATTCode.Left
    lbOtherDate2.Visible = True
    dlpDOther2.Visible = True
    
End Sub

Private Sub SAMUEL_Trans(xEmpNo)
Dim xLDate
    xLDate = dlpTLAYDate(0).Text 'Date
    Call SamuelAuditAdd(xEmpNo, "M", "Enter Leave", "Enter Leave", "", xLDate, xLDate)
End Sub

Private Sub WFC_NGS_Trans(xEmpNo)
Dim xLDate
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    If IsDate(dlpDOther2.Text) Then
        Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", CVDate(dlpDOther2.Text))
    Else
        Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", Null)
    End If
    If IsDate(dlpDOther2.Text) Then
        xLDate = dlpDOther2.Text 'Date
        'Call NGSAuditAdd(xEmpNo, "M", "Enter Leave", "From Date", "", dlpTLAYDate(0).Text, xLDate)
        'Call NGSAuditAdd(xEmpNo, "M", "Enter Leave", "To Date", "", dlpTLAYDate(1).Text, xLDate)
        Call NGSAuditAdd(xEmpNo, "M", "Enter Leave", lStr("Other Date 2"), "", (dlpDOther2.Text), xLDate)
    End If
End Sub

Private Sub SamuelScreenSetup()
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String

comESalInc.Clear
comESalInc.AddItem "Yes"
comESalInc.AddItem "No"
lblTitle(4).Top = 2930
comESalInc.Top = 2930
lblTitle(4).Visible = True
comESalInc.Visible = True

glbEmpDiv = ""
glbEmpAdminBy = ""
glbEmpSection = ""
glbEmpRegion = ""
SQLQ = "SELECT ED_EMPNBR, ED_ADMINBY, ED_DIV, ED_SECTION, ED_REGION FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ADMINBY")) Then glbEmpAdminBy = "" Else glbEmpAdminBy = rsEmpee("ED_ADMINBY")
    If IsNull(rsEmpee("ED_SECTION")) Then glbEmpSection = "" Else glbEmpSection = rsEmpee("ED_SECTION")
    If IsNull(rsEmpee("ED_REGION")) Then glbEmpRegion = "" Else glbEmpRegion = rsEmpee("ED_REGION")
End If
rsEmpee.Close

End Sub

Private Sub CheckReptAuth() 'Ticket #20885 Franks 11/18/2011 for Samuel
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
    xFlag1 = False
    'check if this employee is a Reporting Authority
    If IsReportAuth(glbLEE_ID) Then
        xFlag1 = True
    End If

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    If xFlag1 Then
        xMsg = "This employee has been assigned as a Reporting Authority on other employee files."
        xMsg = xMsg & " Will this LOA affect the Reporting Authority structures?"
        frmMsgYesNoUn.lblMsg.Caption = xMsg
        frmMsgYesNoUn.lblMsg.Alignment = 0
        frmMsgYesNoUn.Show 1
        If glbMsgCustomVal = 1 Or glbMsgCustomVal = 3 Then
            'create a report to show the employee list
            Call CreateEmpList4ReportAuth(glbLEE_ID)
            xSamuleFlag = True
            
'            'show the report - begin
'            Me.vbxCrystal.Reset
'
'            'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'            'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
'            Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
'
'            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
'            If Len(glbstrSelCri) >= 0 Then
'                Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
'            End If
'            'Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & lblEEName & "'"
'            'Ticket #21669 Franks 03/01/2012
'            xMsg = Replace(lblEEName, "'", "''")
'            Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & xMsg & "'"
'
'            Me.vbxCrystal.Connect = RptODBC_SQL
'            Me.vbxCrystal.WindowTitle = "Employee List for Reporting Authority " & lblEEName
'            Me.vbxCrystal.Destination = 0
'            Me.vbxCrystal.Action = 1
'            Me.vbxCrystal.Reset
'            'show the report - end
        End If
    End If
    
End Sub

Private Sub CheckReptDispRpt()
Dim xMsg
        If Not xSamuleFlag Then
            Exit Sub
        End If

        'show the report - begin
        Me.vbxCrystal.Reset
        
        'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
        'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
        Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
        
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
        If Len(glbstrSelCri) >= 0 Then
            Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
        End If
        'Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & lblEEName & "'"
        'Ticket #21669 Franks 03/01/2012
        xMsg = Replace(lblEEName, "'", "''")
        Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & xMsg & "'"
        
        Me.vbxCrystal.Connect = RptODBC_SQL
        Me.vbxCrystal.WindowTitle = "Employee List for Reporting Authority " & lblEEName
        Me.vbxCrystal.Destination = 0
        Me.vbxCrystal.Action = 1
        Me.vbxCrystal.Reset
        'show the report - end
End Sub

Private Sub dlpEndDate_Change() 'Ticket #23920 Franks 07/04/2013
'Call WFCUpdate_Value
xLocID = 0
If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
    xLocID = Data2.Recordset("BM_BENE_ID")
End If
End Sub

Private Sub WFCLastDayForNonNGSEmployee(xEmpNo) 'Ticket #25248 Franks 03/24/2014
'Ticket #25307 Franks 04/08/2014 - use dlpLastDat2 for all employee
'If Not frmWFCBenList.Visible Then
    lblTitle(7).Caption = lStr("Last Day")
    lblTitle(7).Left = lblTitle(3).Left
    If frmWFCBenList.Visible Then
        lblTitle(7).Top = Label1.Top + 420 * 2
        dlpLastDat2.Top = Label1.Top + 420 * 2
    Else
        lblTitle(7).Top = Label1.Top + 420
        dlpLastDat2.Top = Label1.Top + 420
    End If
    dlpLastDat2.Left = clpATTCode.Left
    lblTitle(7).Visible = True
    dlpLastDat2.Visible = True
'End If
End Sub

Private Sub WFCBenListScreen(xEmpNo) 'Ticket #23920 Franks 07/02/2013
Dim rsLEmp As New ADODB.Recordset
Dim rslBen As New ADODB.Recordset
Dim SQLQ As String
    frmWFCBenList.Top = 4440
    frmWFCBenList.Left = 480
    frmWFCBenList.Width = 10575
    frmWFCBenList.Height = 3135
    chkAllDates.Caption = "All Dates"
    'lblTitle(6).Caption = lStr("Last Day")
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsLEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLEmp.EOF Then
        If Not IsNull(rsLEmp("ED_WORKCOUNTRY")) Then
            If rsLEmp("ED_WORKCOUNTRY") = "U.S.A." Then
                If Not IsNull(rsLEmp("ED_VADIM1")) Then
                    If Len(rsLEmp("ED_VADIM1")) > 0 Then
                        Call UpdateBenefitGroup(xEmpNo)
                        
                        Data2.ConnectionString = glbAdoIHRDBW
                        SQLQ = "SELECT * FROM HRBENGRPLIST "
                        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
                        Data2.RecordSource = SQLQ
                        Data2.Refresh

                        frmWFCBenList.Visible = True
                    End If
                Else
                    'Ticket #30607 Franks 09/20/2017 - begin
                    'If ED_VADIM1 is null, then program still needs Data2, even it is blank
                    Data2.ConnectionString = glbAdoIHRDBW
                    SQLQ = "SELECT * FROM HRBENGRPLIST "
                    SQLQ = SQLQ & "WHERE BM_WRKEMP = '****'  "
                    Data2.RecordSource = SQLQ
                    Data2.Refresh
                    'Ticket #30607 Franks 09/20/2017 - end
                End If
            End If
        End If
    End If
End Sub

Private Sub UpdateBenefitGroup(xEmpNo) 'Ticket #23920 Franks 07/02/2013
Dim rsBGMST As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim rsBGEE As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
Dim BelongOldGroup As Boolean
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001W.CommitTrans

    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "ORDER BY BF_BCODE, BF_EDATE "

    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_COMPNO") = "001"
        rsBGTMP("BM_BENEFIT_GROUP") = rsBGMST("BF_GROUP")
        rsBGTMP("BM_BCODE") = rsBGMST("BF_BCODE")
        rsBGTMP("BM_EDATE") = rsBGMST("BF_EDATE")
        rsBGTMP("BM_ENDDATE") = rsBGMST("BF_CEASEDATE") 'New
        rsBGTMP("BM_CHECK") = 1
        rsBGTMP("BM_COVER") = rsBGMST("BF_COVER")
        rsBGTMP("BM_AMT") = rsBGMST("BF_AMT")
        rsBGTMP("BM_PPAMT") = rsBGMST("BF_PPAMT")
        rsBGTMP("BM_UNITCOST") = rsBGMST("BF_UNITCOST")
        rsBGTMP("BM_PCE") = rsBGMST("BF_PCE")
        rsBGTMP("BM_PCC") = rsBGMST("BF_PCC")
        rsBGTMP("BM_ECOST") = rsBGMST("BF_ECOST")
        rsBGTMP("BM_CCOST") = rsBGMST("BF_CCOST")
        rsBGTMP("BM_TCOST") = rsBGMST("BF_TCOST")
        rsBGTMP("BM_MAXDOL") = rsBGMST("BF_MAXDOL")
        rsBGTMP("BM_PREMIUM") = rsBGMST("BF_PREMIUM")
        rsBGTMP("BM_PER") = rsBGMST("BF_PER")
        rsBGTMP("BM_MTHCCOST") = rsBGMST("BF_MTHCCOST")
        rsBGTMP("BM_MTHECOST") = rsBGMST("BF_MTHECOST")
        rsBGTMP("BM_TAXBEN") = rsBGMST("BF_TAXBEN")
        rsBGTMP("BM_SALARYDEPENDANT") = rsBGMST("BF_SALARYDEPENDANT")
        rsBGTMP("BM_MINIMUM") = rsBGMST("BF_MINIMUM")
        rsBGTMP("BM_FACTOR") = rsBGMST("BF_FACTOR")
        rsBGTMP("BM_ROUND") = rsBGMST("BF_ROUND")
        rsBGTMP("BM_MAXIMUM") = rsBGMST("BF_MAXIMUM")
        rsBGTMP("BM_NEXTNEAREST") = rsBGMST("BF_NEXTNEAREST")
        rsBGTMP("BM_TAXAMOUNT") = rsBGMST("BF_TAXAMOUNT")
        rsBGTMP("BM_WAITPERIOD") = rsBGMST("BF_WAITPERIOD")
        rsBGTMP("BM_DWM") = rsBGMST("BF_DWM")
        rsBGTMP("BM_PERORDOLL") = rsBGMST("BF_PERORDOLL")
        rsBGTMP("BM_POLICY") = rsBGMST("BF_POLICY")
        rsBGTMP("BM_RATELEVEL") = rsBGMST("BF_RATELEVEL")
        rsBGTMP("BM_COMMENTS") = rsBGMST("BF_COMMENTS")
        rsBGTMP("BM_PTAX") = rsBGMST("BF_PTAX")
        rsBGTMP("BM_ACTION") = "Add"
        rsBGTMP("BM_WRKEMP") = glbUserID
        
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BF_BCODE") & "' "
        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
        End If
        rsTABL.Close
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
    Call Pause(1)

End Sub

Private Sub WFCNonNGSLastDayUpt(xEmpNo) 'Ticket #25248 Franks 03/24/2014
Dim SQLQ
    'update Last Day
    If dlpLastDat2.Visible Then
        If IsDate(dlpLastDat2.Text) Then
            SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(dlpLastDat2.Text) & " "
            SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
            gdbAdoIhr001.Execute SQLQ
            Call AUDITBENF(xEmpNo, False, , "Y", dlpLastDat2.Text)
        End If
    End If
End Sub

Private Sub WFC_NGSBenEndDateUpt(xEmpNo) 'Ticket #23920 Franks 07/02/2013
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset
Dim rsEmpBN As New ADODB.Recordset
Dim xTemp
Dim xDate1, xDate2
    'Ticket #25307 Franks 04/08/2014 - comment the following codes,
    'use dlpLastDat2 instead of dlpLastDate
    ''''update Last Day
    '''If IsDate(dlpLastDate.Text) Then
    '''    SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(dlpLastDate.Text) & " "
    '''    SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
    '''    gdbAdoIhr001.Execute SQLQ
    '''    Call AUDITBENF(xEmpNo, False, , "Y", dlpLastDate.Text)
    '''End If
    
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBN.EOF
        SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & rsBN("BM_BCODE") & "' "
        If Not IsNull(rsBN("BM_EDATE")) Then SQLQ = SQLQ & "AND BF_EDATE = " & Date_SQL(rsBN("BM_EDATE")) & " "
        If rsEmpBN.State <> 0 Then rsEmpBN.Close
        rsEmpBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpBN.EOF Then
            If IsNull(rsEmpBN("BF_CEASEDATE")) Then xDate1 = CVDate("01/01/1900") Else xDate1 = CVDate(rsEmpBN("BF_CEASEDATE"))
            If IsNull(rsBN("BM_ENDDATE")) Then xDate2 = CVDate("01/01/1900") Else xDate2 = CVDate(rsBN("BM_ENDDATE"))
            rsEmpBN("BF_CEASEDATE") = rsBN("BM_ENDDATE")
            If xNewStatus = "SALC" Then  'Ticket #30446 Franks 08/09/2017
                rsEmpBN("BF_PPAMT") = 0
            End If
            rsEmpBN.Update
            If Not xDate1 = xDate2 Then 'BF_CEASEDATE was changed
                If xDate2 > CVDate("01/01/1900") Then
                    'update hraudit - begin
                    Call AUDITBENF(xEmpNo, False, rsEmpBN)
                    'update hraudit - end
                End If
            End If
        End If
        rsEmpBN.Close
        rsBN.MoveNext
    Loop
    rsBN.Close
End Sub

Private Sub WFCUpdate_Value() 'Ticket #23920 Franks 07/02/2013
Dim SQLQ As String
Dim xID As Long
If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
    
    xID = Data2.Recordset("BM_BENE_ID")
    Data2.Refresh
    'xID = xLocID
    'SQLQ = "BM_BENE_ID = " & xLocID
    If Not IsEmpty(xLocID) Then
        If xLocID > 0 Then
            xID = xLocID
        End If
    End If
    SQLQ = "BM_BENE_ID = " & xID 'xLocID
    Data2.Recordset.Find SQLQ
    
    If IsDate(dlpEndDate.Text) Then
        If Year(dlpEndDate.Text) > 1900 And Year(dlpEndDate.Text) < 2050 Then
            Data2.Recordset("BM_ENDDATE") = dlpEndDate.Text
        Else
            Data2.Recordset("BM_ENDDATE") = Null
        End If
    Else
        Data2.Recordset("BM_ENDDATE") = Null
    End If
    Data2.Recordset.Update
    Data2.Refresh
    DoEvents
    SQLQ = "BM_BENE_ID = " & xID
    Data2.Recordset.Find SQLQ
End If
End Sub

Private Function AUDITBENF(xEmpNo, xlocNewRec As Boolean, Optional rslBen As ADODB.Recordset, Optional xIsWorkDay = "N", Optional xLastDate) 'Ticket #23920 Franks 07/03/2013
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim ACTX
Dim NBCode, NPPAMT, NMTHCOMP, NMTHEMP, NBAMT, NPPE, NPCC, NMAXDOL, NEDate, NCOVER, NTCOST
Dim xTermSEQ
Dim SQLQ As String

On Error GoTo AUDIT_ERR
AUDITBENF = False

If xlocNewRec Then
    ACTX = "A"
Else
    ACTX = "M"
End If

xTermSEQ = 0
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
Else
    SQLQ = "SELECT ED_PT,ED_DIV FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset

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
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_CEASEDATE,AU_LDAY "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

If xIsWorkDay = "N" Then
    NBCode = ""
    NPPAMT = ""
    NMTHCOMP = ""
    NMTHEMP = ""
    NBAMT = ""
    NPPE = ""
    NPCC = ""
    NMAXDOL = ""
    NEDate = ""
    NCOVER = ""
    NTCOST = ""
    NBCode = rslBen("BF_BCODE")
    If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")
    ''If Not IsNull(rslBen("BF_PPAMT")) Then NPPAMT = rslBen("BF_PPAMT")
    ''If Not IsNull(rslBen("BF_MTHCCOST")) Then NMTHCOMP = rslBen("BF_MTHCCOST")
    ''If Not IsNull(rslBen("BF_MTHECOST")) Then NMTHEMP = rslBen("BF_MTHECOST")
    ''If Not IsNull(rslBen("BF_AMT")) Then NBAMT = rslBen("BF_AMT")
    ''If Not IsNull(rslBen("BF_PCC")) Then NPCC = rslBen("BF_PCC")
    ''If Not IsNull(rslBen("BF_PCE")) Then NPPE = rslBen("BF_PCE")
    ''If Not IsNull(rslBen("BF_MAXDOL")) Then NMAXDOL = rslBen("BF_MAXDOL")
    ''If Not IsNull(rslBen("BF_COVER")) Then NCOVER = rslBen("BF_COVER")
    ''If Not IsNull(rslBen("BF_TCOST")) Then NTCOST = rslBen("BF_TCOST")
    ''
    ''If OBCode <> NBCode Then GoTo MODUPD
    '''If OPPE <> NPPE Or OPCC <> NPCC Then GoTo MODUPD
    ''If OPPAMT <> NPPAMT Or OMAXDOL <> NMAXDOL Then GoTo MODUPD
    '''If OMTHCOMP <> NMTHCOMP Or OMTHEMP <> NMTHEMP Then GoTo MODUPD
    ''If OBAMT <> NBAMT Then GoTo MODUPD
    ''If OEDate <> NEDate Then GoTo MODUPD
End If

'GoTo MODNOUPD

'BF_CEASEDATE was changed
MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If xIsWorkDay = "N" Then
    rsTA("AU_BCODE") = NBCode 'clpCode(1).Text
    rsTA("AU_CEASEDATE") = rslBen("BF_CEASEDATE")
    If glbWFC Then 'Ticket #30446 Franks 08/09/2017
        If xNewStatus = "SALC" Then
            If Not IsNull(rslBen("BF_PPAMT")) Then
                If rslBen("BF_PPAMT") = 0 Then
                    rsTA("AU_PPAMT") = 0
                End If
            End If
        End If
    End If
    'If OMTHCOMP <> NMTHCOMP Then rsTA("AU_MTHCCOST") = NMTHCOMP
    'If OMTHEMP <> NMTHEMP Then rsTA("AU_MTHECOST") = NMTHEMP
    'If OTAXBEN <> txtTAXBEN Then rsTA("AU_TAXBEN") = txtTAXBEN
    'If OCOVER <> NCOVER Then rsTA("AU_COVER") = NCOVER
    'If OTCOST <> NTCOST Then rsTA("AU_TCOST") = NTCOST
    'If OPremium <> lblAP Then rsTA("AU_PREMIUM") = lblAP
    'If OPPE <> NPPE Then rsTA("AU_PCE") = NPPE
    'If OPCC <> NPCC Then rsTA("AU_PCC") = NPCC
    'If OPPAMT <> NPPAMT Then
    '    rsTA("AU_PPAMT") = NPPAMT
    '    If IsNumeric(OPPAMT) Then rsTA("AU_OLDPPMT") = Val(OPPAMT)
    'End If
    'If OMAXDOL <> NMAXDOL Then rsTA("AU_MAXDOL") = NMAXDOL
    'If OEDate <> NEDate Then
    '  If IsDate(NEDate) Then
    '      rsTA("AU_EDATE") = CVDate(NEDate)
    '  End If
    'End If
    'If OPER <> txtPer Then rsTA("AU_PER") = txtPer
    'If OBAMT <> NBAMT Then rsTA("AU_BAMT") = NBAMT
    'If OUNITCOST <> medUnitCost Then rsTA("AU_UNITCOST") = IIf(medUnitCost = "", 0, medUnitCost)
    rsTA("AU_LDATE") = Date
    If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
        If CVDate(NEDate) > CVDate(Date) Then
            rsTA("AU_LDATE") = CVDate(NEDate)
        End If
    End If
End If
If xIsWorkDay = "Y" Then
    rsTA("AU_LDAY") = xLastDate
    rsTA("AU_LDATE") = Date
End If
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
Else
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update
rsTA.Close

MODNOUPD:
AUDITBENF = True
Exit Function
AUDIT_ERR:

End Function

Private Sub vbxTrueGrid1_BeforeRowColChange(Cancel As Integer)
'Call WFCUpdate_Value 'Ticket #23920 Franks 07/02/2013
End Sub

'Ticket #23920 Franks 07/03/2013
Private Sub vbxTrueGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
        If IsNull(Data2.Recordset("BM_ENDDATE")) Then
            dlpEndDate.Text = ""
        Else
            dlpEndDate.Text = Data2.Recordset("BM_ENDDATE")
        End If
    End If
End Sub

Private Function WFC_TermEmailSending() As Boolean
'Ticket #24316 - WFC: If Status = 'SALC' then send Termination Email

    WFC_TermEmailSending = False
    
    '' Make sure we have needed info to send email
    'If GetEmpData(glbEmpNbr, "ED_EMAIL") = "" Then ' And Not MDIMain.mnu_File_EmailSetup.Visible Then
    '    Screen.MousePointer = vbDefault
    '    MsgBox GetEmpData(glbEmpNbr, "ED_FNAME") & ", please fill in your email address on the Status/Dates screen, before attempting to terminate an employee.", vbExclamation + vbOKOnly, "Missing Email Address"
    '    Exit Function
    'Else
    '    If Not IsEmailSetup(glbEmpNbr) Then 'MDIMain.mnu_File_EmailSetup.Visible And Not IsEmailSetup(glbEmpNbr) Then  'lost condition afther removing menu items , should check
    '        Screen.MousePointer = vbDefault
    '        MsgBox "You have not been set up for email sending.  Please use the Setup->Security->Email Setup menu option to set up your account for email sending before attempting to put salaried employees on Leave with 'SALC' Employment Status.  Enter a Leave aborted.", vbCritical + vbOKOnly, "No Email Setup Found"
    '        Exit Function
    '    End If
    'End If
    
    ' Send the email
    cmdWFCTermEmail_Click
    
    ' AC - dkostka - 05/03/01 - Added error checking, refuse to terminate if email didn't go through
    If AbortLeave = True Then
        Screen.MousePointer = vbDefault
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(0).Caption = "Enter a Leave Aborted"
        MsgBox "Error sending email.  Enter a Leave aborted.", vbCritical + vbOKOnly, "Error"
        Exit Function
    End If

    'Email sent out successfully
    WFC_TermEmailSending = True
End Function

Private Sub cmdWFCTermEmail_Click()
    Dim MailBody As String
    Dim LocCode As String, LocDesc As String
    Dim xToEmail As String
    
    On Error GoTo ErrorHandler
    
    Load frmSendEmail
    
    'Ticket #18578
    frmSendEmail.txtSubject.Text = "info:HR Termination Notice - " & lblEEName.Caption
    
    MailBody = "The employee below has been terminated." & vbCrLf & vbCrLf
    MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
    MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
    
    ' dkostka - 02/23/01 - Removed Reason from email body, added Location for WFC only.
    'If glbWFC Then
        GetLocation lblEENum.Caption, LocCode, LocDesc
        MailBody = MailBody & "Location: " & LocCode & " - " & LocDesc & vbCrLf
        MailBody = MailBody & "Reporting Authority: " & GetReportingAuthority(lblEENum.Caption) & vbCrLf
        If IsDate(dlpLastDat2.Text) Then
            MailBody = MailBody & "Date: " & dlpLastDat2.Text & vbCrLf & vbCrLf
        End If
    'End If
    frmSendEmail.txtBody.Text = MailBody
    
    ' dkostka - 02/23/2001 - Automated email sending for WFC.
    If glbWFC Then
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(0).Caption = "Sending email..."
        'Franks 05/03/04 Ticket #6105 David Hili wants to change it
        'frmSendEmail.txtTo.Text = "hotline@woodbridgegroup.com"
        frmSendEmail.txtTo.Text = glbWFCTermEmail '"termnotice@woodbridgegroup.com"
        frmSendEmail.Tag = ""
        frmSendEmail.cmdSend_Click
        Do
            DoEvents
        Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
        ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
        '   otherwise refuse to terminate the employee.
        If frmSendEmail.Tag = "DONE" Then
            Unload frmSendEmail
            AbortLeave = False
        Else
            Unload frmSendEmail
            AbortLeave = True
        End If
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(0).FloodType = 1
    'Else
    '    frmSendEmail.Show 1
    End If
    
exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH

End Sub

''Private Function IsEmailSetup(EmpNbr) As Boolean
''    Dim rsEmail As New ADODB.Recordset
''
''    rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
''    If rsEmail.EOF Then
''        IsEmailSetup = False
''    Else
''        IsEmailSetup = True
''    End If
''    rsEmail.Close
''End Function

''Private Function GetLocation(EmpNbr, ByRef LocCode As String, ByRef LocDesc As String)
''    Dim rsEmp As New ADODB.Recordset, rsTABL As New ADODB.Recordset
''
''    rsEmp.Open "SELECT ED_LOC FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
''    If rsEmp.EOF Then
''        LocCode = ""
''        LocDesc = ""
''        rsEmp.Close
''        Exit Function
''    End If
''    If Not IsNull(rsEmp("ED_LOC")) Then
''        LocCode = rsEmp("ED_LOC")
''    Else
''        LocCode = ""
''    End If
''    rsEmp.Close
''
''    rsTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='EDLC' AND TB_KEY='" & LocCode & "'", gdbAdoIhr001
''    If rsTABL.EOF Then
''        LocDesc = ""
''        rsTABL.Close
''        Exit Function
''    End If
''    LocDesc = rsTABL("TB_DESC")
''    rsTABL.Close
''End Function

''Private Function GetReportingAuthority(EmpNbr)
''    Dim rsEmp As New ADODB.Recordset, rsJobHis As New ADODB.Recordset
''    GetReportingAuthority = ""
''    rsJobHis.Open "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr, gdbAdoIhr001
''    If Not rsJobHis.EOF Then
''        If Not IsNull(rsJobHis("JH_REPTAU")) Then
''            If IsNumeric(rsJobHis("JH_REPTAU")) Then
''                rsEmp.Open "SELECT ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & rsJobHis("JH_REPTAU"), gdbAdoIhr001
''                If Not rsEmp.EOF Then
''                    GetReportingAuthority = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
''                End If
''                rsEmp.Close
''            End If
''        End If
''    End If
''    rsJobHis.Close
''End Function

