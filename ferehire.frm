VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEREHIRE 
   Appearance      =   0  'Flat
   Caption         =   "Re-hire An Employee"
   ClientHeight    =   10950
   ClientLeft      =   285
   ClientTop       =   1320
   ClientWidth     =   12285
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10950
   ScaleWidth      =   12285
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   9495
      LargeChange     =   315
      Left            =   12000
      Max             =   1800
      SmallChange     =   315
      TabIndex        =   46
      Top             =   1200
      Width           =   300
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   12285
      _Version        =   65536
      _ExtentX        =   21669
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
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
      Begin VB.Label lblEEID 
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
         TabIndex        =   45
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   43
         Top             =   150
         Width           =   945
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
         TabIndex        =   41
         Top             =   135
         Width           =   720
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   240
         Left            =   4200
         TabIndex        =   42
         Top             =   135
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8160
      Top             =   11640
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ado1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   9960
      Top             =   11640
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ado3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frRehire 
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   240
      TabIndex        =   47
      Top             =   1080
      Width           =   11535
      Begin VB.CommandButton cmdRestoreAll 
         Appearance      =   0  'Flat
         Caption         =   "Restore All"
         Height          =   375
         Left            =   8160
         TabIndex        =   80
         Tag             =   "Recalculate Percentage Change"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtEmpNo 
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
         Left            =   3680
         MaxLength       =   12
         TabIndex        =   20
         Tag             =   "Enter New Employee Number"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Frame fraLinamar 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1035
         Left            =   0
         TabIndex        =   66
         Top             =   4920
         Visible         =   0   'False
         Width           =   8655
         Begin MSMask.MaskEdBox medSIN 
            DataField       =   "ED_SIN"
            DataSource      =   "Data1"
            Height          =   285
            Left            =   3675
            TabIndex        =   24
            Tag             =   "00-Social Insurance Number"
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            ClipMode        =   1
            Appearance      =   0
            PromptInclude   =   0   'False
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "###-###-###"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCountry 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Country will change"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   69
            Top             =   420
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblPROV 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Province will change"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   120
            TabIndex        =   68
            Top             =   780
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lblSIN 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter S.I.N."
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   45
            Visible         =   0   'False
            Width           =   1020
         End
      End
      Begin VB.Frame frmVadim 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   120
         TabIndex        =   59
         Top             =   4530
         Width           =   7965
         Begin INFOHR_Controls.CodeLookup clpPayType 
            Height          =   285
            Left            =   3240
            TabIndex        =   23
            Top             =   30
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   503
         End
         Begin VB.Label lblPaymentType 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Payment Type"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   0
            TabIndex        =   60
            Top             =   75
            Width           =   2160
         End
      End
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
         Left            =   3680
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "01-Employee Payroll ID"
         Top             =   6720
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtEmpType 
         Appearance      =   0  'Flat
         DataSource      =   " "
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
         Left            =   5520
         MaxLength       =   15
         TabIndex        =   58
         Tag             =   "00-Internal Telephone Extension "
         Top             =   7095
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.ComboBox comEmpType 
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
         ItemData        =   "ferehire.frx":0000
         Left            =   3680
         List            =   "ferehire.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Tag             =   "10-Type of Employee "
         Top             =   7080
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame fraBasicInfo 
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   0
         TabIndex        =   48
         Top             =   7800
         Visible         =   0   'False
         Width           =   11535
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   2
            Left            =   3360
            TabIndex        =   37
            Tag             =   "00-New Region"
            Top             =   1800
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDRG"
         End
         Begin INFOHR_Controls.CodeLookup clpDIV1 
            DataField       =   "ED_DIV"
            Height          =   285
            Left            =   3360
            TabIndex        =   31
            Tag             =   "00-New Division"
            Top             =   720
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   1
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "ED_ADMINBY"
            Height          =   285
            Index           =   3
            Left            =   3360
            TabIndex        =   35
            Tag             =   "00-New Administered By"
            Top             =   1440
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDAB"
         End
         Begin INFOHR_Controls.CodeLookup clpDept 
            DataField       =   "ED_DEPTNO"
            Height          =   285
            Left            =   3360
            TabIndex        =   28
            Tag             =   "00-New Department"
            Top             =   0
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   7
            LookupType      =   2
         End
         Begin INFOHR_Controls.CodeLookup clpGLNum 
            DataField       =   "ED_GLNO"
            Height          =   285
            Left            =   3360
            TabIndex        =   30
            Tag             =   "00-New General Ledger - Code"
            Top             =   360
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            MaxLength       =   25
            LookupType      =   3
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            Height          =   285
            Index           =   4
            Left            =   3360
            TabIndex        =   39
            Tag             =   "00-New Section"
            Top             =   2160
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDSE"
         End
         Begin INFOHR_Controls.DateLookup dlpDivEDate 
            DataField       =   "ED_DIVEDATE"
            Height          =   285
            Left            =   8760
            TabIndex        =   32
            Tag             =   "40-Division Effective Date"
            Top             =   720
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.DateLookup dlpDeptEDate 
            DataField       =   "ED_DEPTEDATE"
            Height          =   285
            Left            =   8760
            TabIndex        =   29
            Tag             =   "40-Department Effective Date"
            Top             =   0
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   503
            TextBoxWidth    =   1215
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "ED_LOC"
            Height          =   285
            Index           =   1
            Left            =   3360
            TabIndex        =   33
            Tag             =   "00-New Location - Code"
            Top             =   1080
            Width           =   3195
            _ExtentX        =   5636
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDLC"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "ED_EMP"
            Height          =   285
            Index           =   0
            Left            =   8760
            TabIndex        =   34
            Tag             =   "00-Enter Status Code"
            Top             =   1080
            Visible         =   0   'False
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDEM"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "ED_ORG"
            DataSource      =   " "
            Height          =   285
            Index           =   5
            Left            =   8760
            TabIndex        =   36
            Tag             =   "00-Enter Union Code"
            Top             =   1440
            Visible         =   0   'False
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDOR"
         End
         Begin INFOHR_Controls.CodeLookup clpPT 
            DataField       =   "ED_PT"
            DataSource      =   " "
            Height          =   285
            Left            =   8760
            TabIndex        =   38
            Tag             =   "00-Category Codes"
            Top             =   1800
            Visible         =   0   'False
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDPT"
         End
         Begin VB.Label lblPT 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6720
            TabIndex        =   81
            Top             =   1845
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lblEEStatus 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "New Employment Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6720
            TabIndex        =   79
            Top             =   1125
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lblUnion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "New Union"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6720
            TabIndex        =   78
            Top             =   1485
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New G/L #"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   12
            Left            =   120
            TabIndex        =   57
            Top             =   405
            Width           =   1200
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Department"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   11
            Left            =   120
            TabIndex        =   56
            Top             =   30
            Width           =   1635
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Administered By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   25
            Left            =   120
            TabIndex        =   55
            Top             =   1485
            Width           =   1995
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Region"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   24
            Left            =   120
            TabIndex        =   54
            Top             =   1845
            Width           =   1305
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Location"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   23
            Left            =   120
            TabIndex        =   53
            Top             =   1125
            Width           =   1425
         End
         Begin VB.Label lblTitle 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Enter New Division"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   13
            Left            =   120
            TabIndex        =   52
            Top             =   765
            Width           =   1365
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            Caption         =   "Enter New Section"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   26
            Left            =   120
            TabIndex        =   51
            Top             =   2205
            Width           =   1350
         End
         Begin VB.Label lblDivStart 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Division Effective"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6720
            TabIndex        =   50
            Top             =   765
            Width           =   1245
         End
         Begin VB.Label lblDeptStart 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Department Effective"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6720
            TabIndex        =   49
            Top             =   45
            Width           =   1515
         End
      End
      Begin INFOHR_Controls.DateLookup dlpDOH 
         DataField       =   "ED_DOH"
         Height          =   285
         Left            =   3360
         TabIndex        =   18
         Top             =   3480
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore RSP Contributions Data                         "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Tag             =   "Restore Position/Salary/Performance?"
         Top             =   480
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Position/Salary/Performance                "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Tag             =   "Restore Attendance Data?"
         Top             =   840
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Attendance Data                                   "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Tag             =   "Restore Benefit Data?"
         Top             =   1200
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Benefit Data                                          "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Tag             =   "Restore Health and Safety Data?"
         Top             =   1560
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Health and Safety Data                         "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Tag             =   "Restore WSIB Cost Data?"
         Top             =   1920
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore WSIB Cost Data                                    "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   5
         Left            =   4320
         TabIndex        =   17
         Tag             =   "Delete Termination Record?"
         Top             =   3000
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Delete Termination Record                            "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   0
         Tag             =   "Restore Employee Training Information?"
         Top             =   120
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Training                                                 "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   12
         Left            =   4320
         TabIndex        =   14
         Tag             =   "Restore Dollar Entitlements  Data ?"
         Top             =   1920
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Dollar Entitlements  Data                  "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   11
         Left            =   4320
         TabIndex        =   13
         Tag             =   "Restore Trade  Data ?"
         Top             =   1560
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Trade  Data                                      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   10
         Left            =   4320
         TabIndex        =   12
         Tag             =   "Restore Skills Data?"
         Top             =   1200
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Skills Data                                        "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   9
         Left            =   4320
         TabIndex        =   11
         Tag             =   "Restore Formal Education Data ?"
         Top             =   840
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Formal Education Data                     "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   8
         Left            =   4320
         TabIndex        =   10
         Tag             =   "Restore Other Earnings Data?"
         Top             =   480
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Other Earnings Data                        "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   7
         Left            =   4320
         TabIndex        =   9
         Tag             =   "Restore Comments Data ?"
         Top             =   120
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Comments Data                               "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   6
         Tag             =   "Restore Counselling Data        "
         Top             =   2280
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Counselling Data                                   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpLTHire 
         DataField       =   "ED_LTHIRE"
         Height          =   285
         Left            =   3360
         TabIndex        =   19
         Top             =   3840
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   15
         Left            =   4320
         TabIndex        =   16
         Top             =   2640
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Attachment Data                              "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   7
         Tag             =   "Restore Attendance Data?"
         Top             =   2640
         Width           =   3735
         _Version        =   65536
         _ExtentX        =   6588
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Follow-ups Data                                   "
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpDOther1 
         DataSource      =   " "
         Height          =   285
         Left            =   3360
         TabIndex        =   27
         Tag             =   "40-Other Date 2"
         Top             =   7440
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin Threed.SSCheck chkRestore 
         Height          =   195
         Index           =   17
         Left            =   4320
         TabIndex        =   15
         Tag             =   "Restore Hourly Entitlements  Data ?"
         Top             =   2280
         Width           =   3495
         _Version        =   65536
         _ExtentX        =   6165
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Restore Hourly Entitlements  Data                 "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin VB.Frame frmlinamar 
         BorderStyle     =   0  'None
         Caption         =   "Enter New Employee Number"
         Height          =   615
         Left            =   2520
         TabIndex        =   61
         Top             =   4200
         Visible         =   0   'False
         Width           =   3495
         Begin VB.TextBox txtEmpID 
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
            Left            =   1150
            MaxLength       =   6
            TabIndex        =   22
            Tag             =   "01-Employee ID in the Division"
            Top             =   330
            Width           =   945
         End
         Begin INFOHR_Controls.CodeLookup clpDIV 
            Height          =   285
            Left            =   840
            TabIndex        =   21
            Tag             =   "01-Division"
            Top             =   0
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "n/a"
            LookupType      =   1
         End
         Begin VB.Label lblEENumNew 
            AutoSize        =   -1  'True
            Caption         =   "lblEENumNew"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2250
            TabIndex        =   65
            Top             =   375
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblTitle 
            Caption         =   "Employee ID"
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
            Index           =   2
            Left            =   60
            TabIndex        =   63
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label lblTitle 
            Caption         =   "Facility"
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
            Index           =   3
            Left            =   60
            TabIndex        =   64
            Top             =   45
            Width           =   1035
         End
         Begin VB.Label lblEEIDNew 
            Caption         =   "lblEEIDNew"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2460
            TabIndex        =   62
            Top             =   180
            Visible         =   0   'False
            Width           =   915
         End
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New ""Original Hire Date"""
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   77
         Top             =   3525
         Width           =   2880
      End
      Begin VB.Label lblEmpno 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New Employee Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   76
         Top             =   4245
         Width           =   2475
      End
      Begin VB.Label lblEmpExist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number Already Exist"
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
         Height          =   240
         Left            =   5160
         TabIndex        =   75
         Top             =   4215
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New ""Last Hire Date"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   74
         Top             =   3885
         Width           =   2880
      End
      Begin VB.Label lblPayIDExist 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Number Already Exist"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4680
         TabIndex        =   73
         Top             =   6765
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll ID"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   72
         Top             =   6765
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblEEType 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   71
         Top             =   7140
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label lbOtherDate1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Date 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   70
         Top             =   7485
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Employee Master will be restored"
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
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   600
      Width           =   9015
   End
End
Attribute VB_Name = "frmEREHIRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oSENDTE
Dim rsDATA2 As New ADODB.Recordset
Dim fglbNew
Dim xUpdateable
Dim oDOH
Dim oFday
Dim oPayrollID
Dim oDOT
Dim OldUnion, OldDept, OldDiv, OldLoc, OldAdmin, OldRegion, OldSection, OldEmpStatus 'WFC
Dim locPlantCode 'WFC
Dim MailBody
Dim locNewHire As Boolean
Dim SaveBGroup, NewBGroup, NewPayGroup, NewNGSSub 'Ticket #23247 Franks 09/17/2013
Dim xHRSoftUpt As Boolean
Dim SaveGLNo
Dim IsWFCNGSEmployee As Boolean

Private Function ChkInput()
Dim Msg$, Response%
ChkInput = False

If Len(dlpDOH) > 0 Then
    If Not IsDate(dlpDOH) Then
        MsgBox "Invalid Date of Hire"
        dlpDOH.SetFocus
        Exit Function
    End If
Else
    MsgBox lStr("New Original Hire Date is a required field.")
    dlpDOH.SetFocus
    Exit Function
End If

If lblEmpExist.Visible = True Then
    If Len(txtEmpNo) = 0 Then
        MsgBox "Employee Number Missing"
        txtEmpNo.SetFocus
        Exit Function
    End If
    If Not IsNumeric(txtEmpNo) And Not glbLinamar Then
        MsgBox "Invalid Employee Number"
        txtEmpNo.SetFocus
        Exit Function
    End If
    If Not EmpNoExist(getEmpnbr(txtEmpNo)) Then
        MsgBox "Employee Number Already Exist"
        txtEmpNo.SetFocus
        Exit Function
    End If
End If
If glbLinamar Then
    If frmlinamar.Visible Then
        If clpDiv.Caption = "Unassigned" Or Len(clpDiv) <> 3 Or Not IsNumeric(clpDiv) Then
            MsgBox lStr("Invalid Division")
            clpDiv.SetFocus
            Exit Function
        End If
        If Len(txtEmpID) = 0 Then
            MsgBox "Employee ID is a required field"
            txtEmpID.SetFocus
            Exit Function
        Else
            If Not IsNumeric(txtEmpID) Then
                MsgBox "Invalid Employee ID"
                txtEmpID.SetFocus
                Exit Function
            Else
                If Val(txtEmpID) = 0 Then
                    MsgBox "Invalid Employee ID"
                    txtEmpID.SetFocus
                    Exit Function
                End If
            End If
        End If
        If Not EmpNoExist(getEmpnbr(lblEENumNew)) Then
            MsgBox "Employee Number Already Exist"
            txtEmpID.SetFocus
            Exit Function
        End If
    End If
Else
    'If lblEmpExist.Visible = True Then
        If Len(txtEmpNo) = 0 Then
            MsgBox "Employee Number Missing"
            txtEmpNo.SetFocus
            Exit Function
        End If
        If Not IsNumeric(txtEmpNo) And Not glbLinamar Then
            MsgBox "Invalid Employee Number"
            txtEmpNo.SetFocus
            Exit Function
        End If
        If Not EmpNoExist(getEmpnbr(txtEmpNo)) Then
            MsgBox "Employee Number Already Exist"
            txtEmpNo.SetFocus
            Exit Function
        End If
    'End If
End If
If glbLinamar And medSIN.Visible Then
    If gSec_Show_SIN_SSN Then
        If Len(medSIN.Text) = 0 Then
            MsgBox "SIN Number is required field"
            medSIN.SetFocus
            Exit Function
        Else
            If Not SIN_chk(medSIN.Text) Then
                MsgBox "Invalid SIN" & IIf(glbLinamar, "", "- if Unassigned set to 999-999-999")
                medSIN.SetFocus
                Exit Function
            End If
        End If
    End If
    If CheckSINSSNGen(medSIN, "SIN") Then
        Msg$ = "Duplicate SIN number found. "
        Msg$ = Msg$ & Chr(10) & "Are you sure you wish to accept it?"
        Msg$ = Msg$ & Chr(10) & "Press Yes to accept or No to edit"
        Response% = MsgBox(Msg, MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2, "")    ' Get user response.
        If Response% = IDNO Then
            medSIN.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #27606 Franks 10/02/2015 - remove this logic for them
'If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/24/2014 Franks
'    If CheckSINSSNGen(medSIN, "SIN") Then
'        Msg$ = "Duplicate SIN found. "
'        Msg$ = Msg$ & Chr(10) & "This employee cannot be rehired"
'        MsgBox Msg$
'        Exit Function
'    End If
'End If

If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then
    If Len(txtPayrollID.Text) = 0 Then
        MsgBox "Payroll ID is required field"
        txtPayrollID.SetFocus
        Exit Function
    Else
           If glbWFC Then
                If glbPlantCode = "GREN" And Len(txtPayrollID.Text) <> 6 Then
                    MsgBox "Invalid format for Payroll ID. Format must be ######"
                    txtPayrollID.SetFocus
                    Exit Function
                End If
            End If
    End If
    
    Call PayIDExist(txtPayrollID)
    If lblPayIDExist.Visible Then
        Msg$ = "Payroll ID " & txtPayrollID & " already exists "
        'Msg$ = Msg$ & Chr(10) & " A NEW Payroll ID is required"
        MsgBox Msg$
        txtPayrollID.SetFocus
        Exit Function
    End If
        
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire
    If dlpLTHire.Visible Then
        If Len(dlpLTHire.Text) = 0 Then
            MsgBox lStr("Last Hire") & " is required field"
            dlpLTHire.SetFocus
            Exit Function
        End If
    End If
End If

'Ticket #19310 - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    If Len(Trim(clpDept.Text)) = 0 Then
        MsgBox lStr("Department") & " is a required field"
        clpDept.SetFocus
        Exit Function
    Else
        If Not clpDept.ListChecker Then Exit Function
    End If
    
    If Len(Trim(dlpDeptEDate)) = 0 Then
        MsgBox lStr("Department Effective") & " is a required field"
        dlpDeptEDate.SetFocus
        Exit Function
    ElseIf Not IsDate(dlpDeptEDate) Then
        MsgBox lStr("Department Effective") & " is a invalid date"
        dlpDeptEDate.SetFocus
        Exit Function
    End If
    
    'They do not want this field to be mandatory
    'If Len(Trim(clpGLNum.Text)) = 0 Then
    '    MsgBox lStr("G/L #") & " is a required field"
    '    clpGLNum.SetFocus
    '    Exit Function
    'Else
        If clpGLNum.Caption = "Unassigned" Then
            MsgBox lStr("G/L #") & " is invalid"
            clpGLNum.SetFocus
            Exit Function
        End If
    'End If
    
    If Len(Trim(clpDiv1.Text)) = 0 Then
        MsgBox lStr("Division") & " is a required field"
        clpDiv1.SetFocus
        Exit Function
    Else
        If Not clpDiv1.ListChecker Then Exit Function
    End If
    
    'They do not want this field to be mandatory
    'If Len(Trim(dlpDivEDate)) = 0 Then
    '    MsgBox lStr("Division Effective") & " is a required field"
    '    dlpDivEDate.SetFocus
    '    Exit Function
    'ElseIf Not IsDate(dlpDivEDate) Then
        If Len(dlpDivEDate) > 0 Then
            If Not IsDate(dlpDivEDate) Then
                MsgBox lStr("Division Effective") & " is a invalid date"
                dlpDivEDate.SetFocus
                Exit Function
            End If
        End If
    'End If
        
    If Len(Trim(clpCode(1).Text)) = 0 Then
        MsgBox lStr("Location") & " is a required field"
        clpCode(1).SetFocus
        Exit Function
    Else
        If Not clpCode(1).ListChecker Then Exit Function
    End If
    
    If Len(Trim(clpCode(3).Text)) = 0 Then
        MsgBox lStr("Administered By") & " is a required field"
        clpCode(3).SetFocus
        Exit Function
    Else
        If Not clpCode(3).ListChecker Then Exit Function
    End If
    
    If Len(Trim(clpCode(2).Text)) = 0 Then
        MsgBox lStr("Region") & " is a required field"
        clpCode(2).SetFocus
        Exit Function
    Else
        If Not clpCode(2).ListChecker Then Exit Function
    End If
    
    If Len(Trim(clpCode(4).Text)) = 0 Then
        MsgBox lStr("Section") & " is a required field"
        clpCode(4).SetFocus
        Exit Function
    Else
        If Not clpCode(4).ListChecker Then Exit Function
    End If
Else
    'If Len(Trim(clpDept.Text)) = 0 Then
    '    MsgBox lStr("Department") & " is a required field"
    '    clpDept.SetFocus
    '    Exit Function
    'Else
        If Not clpDept.ListChecker Then Exit Function
    'End If
    
    'If Len(Trim(dlpDeptEDate)) = 0 Then
    '    MsgBox lStr("Department Effective") & " is a required field"
    '    dlpDeptEDate.SetFocus
    '    Exit Function
    'ElseIf Not IsDate(dlpDeptEDate) Then
    If Len(dlpDeptEDate) > 0 Then
        If Not IsDate(dlpDeptEDate) Then
            MsgBox lStr("Department Effective") & " is a invalid date"
            dlpDeptEDate.SetFocus
            Exit Function
        End If
    End If
    
    'If Len(Trim(clpGLNum.Text)) = 0 Then
    '    MsgBox lStr("G/L #") & " is a required field"
    '    clpGLNum.SetFocus
    '    Exit Function
    'Else
        If clpGLNum.Caption = "Unassigned" Then
            MsgBox lStr("G/L #") & " is invalid"
            clpGLNum.SetFocus
            Exit Function
        End If
    'End If
    
    'If Len(Trim(clpDIV1.Text)) = 0 Then
    '    MsgBox lStr("Division") & " is a required field"
    '    clpDIV1.SetFocus
    '    Exit Function
    'Else
        If Not clpDiv1.ListChecker Then Exit Function
    'End If
    
    'If Len(Trim(dlpDivEDate)) = 0 Then
    '    MsgBox lStr("Division Effective") & " is a required field"
    '    dlpDivEDate.SetFocus
    '    Exit Function
    'ElseIf Not IsDate(dlpDivEDate) Then
    If Len(dlpDivEDate) > 0 Then
        If Not IsDate(dlpDivEDate) Then
            MsgBox lStr("Division Effective") & " is a invalid date"
            dlpDivEDate.SetFocus
            Exit Function
        End If
    End If
        
    'If Len(Trim(clpCode(1).Text)) = 0 Then
    '    MsgBox lStr("Location") & " is a required field"
    '    clpCode(1).SetFocus
    '    Exit Function
    'Else
        If Not clpCode(1).ListChecker Then Exit Function
    'End If
    
    'If Len(Trim(clpCode(3).Text)) = 0 Then
    '    MsgBox lStr("Administered By") & " is a required field"
    '    clpCode(3).SetFocus
    '    Exit Function
    'Else
        If Not clpCode(3).ListChecker Then Exit Function
    'End If
    
    'If Len(Trim(clpCode(2).Text)) = 0 Then
    '    MsgBox lStr("Region") & " is a required field"
    '    clpCode(2).SetFocus
    '    Exit Function
    'Else
        If Not clpCode(2).ListChecker Then Exit Function
    'End If
    
    'If Len(Trim(clpCode(4).Text)) = 0 Then
    '    MsgBox lStr("Section") & " is a required field"
    '    clpCode(4).SetFocus
    '    Exit Function
    'Else
        If Not clpCode(4).ListChecker Then Exit Function
    'End If
    
End If

'Ticket #23820 Franks 05/29/2013 - begin
If lblEEStatus.Visible Then
    If Len(Trim(clpCode(0).Text)) = 0 Then
        If glbWFC Then
            MsgBox lblEEStatus.Caption & " is a required field"
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    If Not clpCode(0).ListChecker Then Exit Function
    If glbWFC Then '
        If IsLOATypeCode(clpCode(0).Text) Then
            MsgBox "Employment Status cannot be a LOA status"
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
End If
If lblUnion.Visible Then
    If Len(Trim(clpCode(5).Text)) = 0 Then
        If glbWFC Then
            MsgBox lblUnion.Caption & " is a required field"
            clpCode(5).SetFocus
            Exit Function
        End If
    End If
    If Not clpCode(5).ListChecker Then Exit Function
End If
If lblPT.Visible And lblPT.FontBold Then 'Ticket #25562 Franks 06/17/2014
    If Len(Trim(clpPT.Text)) = 0 Then
        MsgBox lblPT.Caption & " is a required field"
        clpPT.SetFocus
        Exit Function
    End If
    If Not clpCode(5).ListChecker Then Exit Function
End If
'Ticket #23820 Franks 05/29/2013 - end
If glbWFC Then
    If Len(Trim(clpDiv1.Text)) = 0 Then 'Ticket #24582 Franks 11/11/2013
            MsgBox lStr("Division") & " is a required field"
            clpDiv1.SetFocus
            Exit Function
    End If
    If dlpDOther1.Visible Then 'Ticket #24652 Franks 12/02/2013
        If Len(Trim(dlpDOther1.Text)) = 0 Then 'Ticket #24582 Franks 11/11/2013
            If clpPT.Text = "FT" Then 'Ticket #25562 Franks 06/17/2014
                MsgBox lbOtherDate1 & " is a required field"
                dlpDOther1.SetFocus
                Exit Function
            End If
        Else
            If Not IsDate(dlpDOther1.Text) Then
                MsgBox "Invalid " & lbOtherDate1 & "."
                dlpDOther1.SetFocus
                Exit Function
            End If
        End If
    End If
End If
    
ChkInput = True

End Function

Private Sub PayIDExist(xPayID)
Dim SQLQ
Dim rsEmp As New ADODB.Recordset
If xPayID = "" Then xPayID = 0
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP"
SQLQ = SQLQ & " WHERE HREMP.ED_PAYROLL_ID = '" & xPayID & "' "
If glbWFC Then
    SQLQ = SQLQ & " AND HREMP.ED_SECTION = '" & locPlantCode & "' "
End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsEmp.EOF And rsEmp.BOF Then
    lblPayIDExist.Caption = ""
    lblPayIDExist.Visible = False
'    lblPayIDExist.Caption = ""
'    lblPayIDExist.Visible = False
'    rsEMP.Close
'    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP"
'    SQLQ = SQLQ & " WHERE Term_HREMP.ED_PAYROLL_ID = '" & xPayID & "' "
'    rsEMP.Open SQLQ, gdbAdoIhr001X, adOpenStatic
'    If Not (rsEMP.EOF And rsEMP.BOF) Then
'        lblPayIDExist.Caption = " Payroll ID " & xPayID & " already exists " '- A NEW Payroll ID is required"
'        lblPayIDExist.Visible = True
'    End If
Else
    lblPayIDExist.Caption = " Payroll ID " & xPayID & " already exists " '- A NEW Payroll ID is required"
    lblPayIDExist.Visible = True
End If
rsEmp.Close


End Sub

Private Function getDOT(xTERM_Seq) 'Ticket #19231
Dim SQLQ, rsTerm As New ADODB.Recordset
Dim retVal
    retVal = ""
    SQLQ = "SELECT Term_DOT,TERM_SEQ FROM TERM_HRTRMEMP "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq & " "
    rsTerm.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    If Not rsTerm.EOF Then
        If Not IsNull(rsTerm("Term_DOT")) Then
            retVal = rsTerm("Term_DOT")
        End If
    End If
    getDOT = retVal
End Function

Private Function isTOUT() As Boolean
Dim SQLQ, rsTerm As New ADODB.Recordset

isTOUT = False
SQLQ = "SELECT TERM_REASON FROM TERM_HRTRMEMP"
SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq & ";"
rsTerm.Open SQLQ, gdbAdoIhr001X
If rsTerm.EOF Then Exit Function

If rsTerm("TERM_REASON") = "TOUT" Then isTOUT = True

End Function

Private Sub chkRestore_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Sub cmdClose_Click()

On Error GoTo err_Unload

glbTERM_ID = 0
glbTERM_Seq = 0
glbOnTop = ""

If glbWFC And glbCandidate > 0 And xHRSoftUpt Then
    'Ticket #24451 Franks 10/09/2013
    '- fix "After the regular REHIRE function finished, the NEW Employee box popped up?"
Else
    Call MDIMain.mmnu_Active_Click
End If

Unload Me

Exit Sub

err_Unload:
Unload Me
Resume Next
Unload Me

End Sub


'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(Me.ActiveControl) '19Aug99 js
'End Sub
Function CheckSINSSNGen(xSINSNN, TypeFlag)
Dim RsSIN As New ADODB.Recordset
Dim SQLQ
If Not glbLinamar Then If xSINSNN = "999999999" Then Exit Function
    CheckSINSSNGen = False
    SQLQ = "SELECT ED_EMPNBR,ED_SIN,ED_SSN FROM HREMP "
    If TypeFlag = "SIN" Then
        SQLQ = SQLQ & "WHERE ED_SIN = '" & xSINSNN & "' "
    Else
        If IsNull(xSINSNN) Or Len(xSINSNN) = 0 Then Exit Function
        SQLQ = SQLQ & "WHERE ED_SSN = '" & xSINSNN & "' "
    End If
    RsSIN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not RsSIN.EOF Then
        CheckSINSSNGen = True
    End If
    RsSIN.Close
    
End Function

Private Sub CountEmpNbr()
lblEENumNew.Visible = True
If Len(clpDiv) = 3 And Val(txtEmpID) > 0 Then
    lblEENumNew = Format(clpDiv, "000") & "-" & Val(txtEmpID)
    lblEEIDNew = Val(txtEmpID) & Format(clpDiv, "000")
Else
    lblEENumNew = ""
End If
End Sub

Sub cmdOK_Click()
Dim Msg$, Title$, DgDef As Variant, Response%, EID&, X%
Dim intLastR%, TermDate$, SQLQ, xlen, SEQID&
Dim xWDate, xVAC, xSICK, xFDate, xTDate, xWDateS, xfdateS, xtdateS
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset

If lblEENum = 0 Then Exit Sub

If Not ChkInput() Then Exit Sub

'Msg$ = Msg$ & Chr(10) & "Are you sure you want to reinstate "
Msg$ = Msg$ & "Are you sure you want to reinstate "
'Msg$ = Msg$ & Chr(10) & "this employee ?"
Msg$ = Msg$ & "this employee ?"
'Msg$ = Msg$ & Chr(10) & "Make sure no other info:HR Window "
Msg$ = Msg$ & Chr(10) & Chr(10) & "Note: Make sure no other info:HR Window "
'Msg$ = Msg$ & Chr(10) & "is open with this employee information showing"
Msg$ = Msg$ & "is open with this employee information showing."

Title$ = "Reinstate Employee"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
Else
    'Ticket #20726 - Check for Employee License if not exceeding
    If Not modECountChk() Then
        MsgBox "You have reached the maximum number of employees for your license. You cannot reinstate this employee." & vbCrLf & vbCrLf & "Please contact HR Systems Strategies Inc.", vbExclamation, "info:HR - Employee License"
        Exit Sub
    End If
End If

DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg = "Are you sure you want to restore Attendance Data?"

If chkRestore(1).Value = True Then
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        'Exit Sub   'Ticket #13871 - Jerry said to proceed with the rehire because the user has already said Yes to the prompt above.
        chkRestore(1).Value = False
        MsgBox "Attendance will not be restored", vbOKOnly, "Restore Attendance"
    End If
End If
If glbCompSerial = "S/N - 2372W" Then 'Town of Bradford West Gwillimbury
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg = "Will this be a new employee in the USTI Payroll?"
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        locNewHire = False
    Else
        locNewHire = True
    End If
End If
    
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodType = 1

If Len(txtEmpNo) = 0 Then
    EID& = lblEENum
Else
    EID& = getEmpnbr(txtEmpNo)
End If

SEQID& = CLng(glbTERM_Seq)
''Ticket #22819 Franks 11/23/2012
'If glbLinamar Then
'    rsDATA2("ED_COUNTRY") = "CANADA"
'    rsDATA2("ED_PROV") = "ON"
'    rsDATA2("ED_SIN") = medSIN
'End If
If glbVadim Then
    If Vadim_PayType_field <> "" Then
        rsDATA2(Vadim_PayType_field) = clpPayType
    End If
End If

'Ticket #20270 Franks 05/05/2011
If IsNull(rsDATA2("ED_WORKCOUNTRY")) Then
    glbEmpCountry = ""
Else
    glbEmpCountry = rsDATA2("ED_WORKCOUNTRY")
End If

'04/05/2010 By Frank - Jerry asked to update Employee Federal and Prov. Tax Exemption using Company Master amounts
'SERVER is down, can't create ticket for this.
Call UptTaxExampt

'Jaddy Removed this because this not make sence
'rsDATA2("ED_DOH") = dlpDOH
'rsDATA2("ED_SENDTE") = dlpDOH
rsDATA2.Update
If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/24/2014 Franks
    'move the check to ChkInput
Else
    If CheckSINSSNGen(rsDATA2("ED_SIN"), "SIN") Then
        Msg$ = "Duplicate SIN number found"
        Msg$ = Msg$ & Chr(10) & "To change SIN, goto the Demographics screen"
        MsgBox Msg$
    End If
    If CheckSINSSNGen(rsDATA2("ED_SSN"), "SSN") Then
        Msg$ = "Duplicate SSN number found"
        Msg$ = Msg$ & Chr(10) & "To change SSN, goto the Demographics screen"
        MsgBox Msg$
    End If
    'Add by Franks Feb 04,2002
End If
If Not modReinMove(EID, SEQID, TermDate) Then Exit Sub
DoEvents

If glbWFC Then 'Ticket #19266 Franks  12/02/2010
    If dlpDOther1.Visible Then
        Call WFC_NGS_Trans(EID&)
    End If
End If

If glbCompSerial = "S/N - 2439W" Then   'OK Tire - Ticket #21518 Franks 07/06/2012
    Call AUDIT_GWL_TRANS(EID&)
End If

Screen.MousePointer = DEFAULT

If gsEMAIL_ONREHIRE Then
    'If NewHireForms.count > 0 Then 'new hire   'Hemu commented it because it does not make sense and it will never be true
        Screen.MousePointer = DEFAULT
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352
            MailBody = GetEmailBodyForSamuel(EID)
            MailBody = MailBody & "has been rehired." & vbCrLf & vbCrLf
            MailBody = MailBody & "Old Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "New Employee #: " & EID& & vbCrLf
            Call EmailSendingForSamuel
        Else
            MailBody = "This employee has been rehired." & vbCrLf & vbCrLf
            MailBody = MailBody & "Old Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "New Employee #: " & EID& & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            Call imgEmail_Click
        End If
    'End If
End If
Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodPercent = 100
If Data1.Recordset.RecordCount <= 1 Then intLastR% = True

If chkRestore(5).Value = True Then
    glbRest = True
    If Not modNukeEETerm(SEQID&) Then MsgBox "Employee remains in Termination file."
Else
    rsTB.Open "Term_HRTRMEMP", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsTB.Find "TERM_SEQ = " & SEQID
    If Not rsTB.EOF Then
        rsTB("Term_DOR") = Format(Now, "SHORT DATE")
        rsTB.Update
    End If
    rsTB.Close
End If
glbRest = False

''Since V7.6
'Ticket #22392 Franks 08/03/2012
'comment out EntReCalc since it caused ED_DHRS blank problem because there is no current position at this moment
'Call EntReCalc("ED_EMPNBR=" & EID&)

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    GoTo Bypass
End If

Data3.Refresh
SQLQ = "Select * from HREMP"
SQLQ = SQLQ & " where HREMP.ED_EMPNBR = " & glbTERM_ID
Data1.RecordSource = SQLQ
Data1.Refresh

If Data1.Recordset.BOF And Data1.Recordset.EOF Then GoTo Bypass


'Data1.Recordset.Edit

'Data1.Recordset("ED_ENTOPT") = glbEntOutStanding$
'Data1.Recordset("ED_ENTOPTS") = glbEntOutStandingS$
'Data1.Recordset("ED_EFDATE") = Null
'Data1.Recordset("ED_ETDATE") = Null
'Data1.Recordset("ED_EFDATES") = Null
'Data1.Recordset("ED_ETDATES") = Null
'Data1.Recordset("ED_PVAC") = 0
'Data1.Recordset("ED_VAC") = 0
'Data1.Recordset("ED_PSICK") = 0
'Data1.Recordset("ED_SICK") = 0
'Data1.Recordset("ED_VACT") = 0
'Data1.Recordset("ED_SICKT") = 0
'
'xWDate = ""
'xFDate = ""
'xTDate = ""
'xWDateS = ""
'xfdateS = ""
'xtdateS = ""
'
''*glbEntOutStanding' '$*
'If glbEntOutStanding$ = "2" Then xWDate = Data1.Recordset.Fields("ED_DOH")
'If glbEntOutStanding$ = "3" Then xWDate = Data1.Recordset.Fields("ED_SENDTE")
'If glbEntOutStanding$ = "4" Then xWDate = Data1.Recordset.Fields("ED_LTHIRE")
'If glbEntOutStanding$ = "5" Then xWDate = Data1.Recordset.Fields("ED_USRDAT1")
'If glbEntOutStanding$ = "6" Then xWDate = Data1.Recordset.Fields("ED_UNION")
'If glbEntOutStanding$ = "1" Then xWDate = glbCompEdFrom
'
''*glbEntOutStanding'S'$'*
'If glbEntOutStandingS$ = "2" Then xWDateS = Data1.Recordset.Fields("ED_DOH")
'If glbEntOutStandingS$ = "3" Then xWDateS = Data1.Recordset.Fields("ED_SENDTE")
'If glbEntOutStandingS$ = "4" Then xWDateS = Data1.Recordset.Fields("ED_LTHIRE")
'If glbEntOutStandingS$ = "5" Then xWDateS = Data1.Recordset.Fields("ED_USRDAT1")
'If glbEntOutStandingS$ = "6" Then xWDateS = Data1.Recordset.Fields("ED_UNION")
'If glbEntOutStandingS$ = "1" Then xWDateS = glbCompEdFromS
'
'If IsDate(xWDate) Then
'    If glbEntOutStanding$ = "1" Then
'        xFDate = glbCompEdFrom
'        xTDate = glbCompEdTo
'    Else
'        xlen = InStr(4, xWDate, "/")
'        xWDate = Left(xWDate, xlen) & Year(Now)
'        If DateValue(xWDate) <= Now Then
'            xFDate = xWDate
'            xWDate = DateAdd("d", 365, xWDate)
'            xTDate = xWDate
'        Else
'            xFDate = xWDate
'            xWDate = DateAdd("d", -365, xWDate)
'            xTDate = xWDate
'        End If
'    End If
'End If
'
'If IsDate(xWDateS) Then
'    If glbEntOutStandingS$ = "1" Then
'        xfdateS = glbCompEdFromS
'        xtdateS = glbCompEdToS
'    Else
'        xlen = InStr(4, xWDateS, "/")
'        xWDateS = Left(xWDateS, xlen) & Year(Now)
'        If DateValue(xWDateS) <= Now Then
'            xfdateS = xWDateS
'            xWDateS = DateAdd("d", 365, xWDateS)
'            xtdateS = xWDateS
'        Else
'            xfdateS = xWDateS
'            xWDateS = DateAdd("d", -365, xWDateS)
'            xtdateS = xWDateS
'        End If
'    End If
'End If
'
'xVAC = 0
'xSICK = 0
'
'
'Do Until Data3.Recordset.EOF
'    If Not Data3.Recordset.EOF Then
'        If IsDate(xWDate) Then
'            If Data3.Recordset.Fields("AD_DOA") >= DateValue(xFDate) And Data3.Recordset.Fields("AD_DOA") < DateValue(xTDate) Then
'                If Left(Data3.Recordset.Fields("AD_REASON"), 3) = "VAC" Then
'                    xVAC = xVAC + Data3.Recordset.Fields("AD_HRS")
'                End If
'            End If
'        End If
'        If IsDate(xWDateS) Then
'            If Data3.Recordset.Fields("AD_DOA") >= DateValue(xfdateS) And Data3.Recordset.Fields("AD_DOA") < DateValue(xtdateS) Then
'                If Left(Data3.Recordset.Fields("AD_REASON"), 3) = "SIC" Then
'                    xSICK = xSICK + Data3.Recordset.Fields("AD_HRS")
'                End If
'            End If
'        End If
'    End If
'    Data3.Recordset.MoveNext
'Loop
'
'If IsDate(xWDate) Then
'    Data1.Recordset.Fields("ED_EFDATE") = xFDate
'    Data1.Recordset.Fields("ED_ETDATE") = xTDate
'End If
'
'If IsDate(xWDateS) Then
'    Data1.Recordset.Fields("ED_EFDATES") = xfdateS
'    Data1.Recordset.Fields("ED_ETDATES") = xtdateS
'End If
'
'Data1.Recordset.Fields("ED_VACT") = xVAC
'Data1.Recordset.Fields("ED_SICKT") = xSICK
Data1.Recordset.Resync

Bypass:

'Hemu - Ticket #9802 - Jerry said if users has selected to reinstate Pos/Sal/Perf then inform them
'to reassign Current records as we have turned off the current flags
If glbWFC Then
    'Ticket #24695 Franks 11/28/2013 - wfc not need this
Else
    If chkRestore(0).Value = True Then
        MsgBox "Make sure you set the Current Position, Salary and Performance record for this employee.", vbInformation, "Reinstate Employee"
    End If
End If
'Hemu

For X% = 0 To 12
  chkRestore(X%).Value = False
Next X%
If glbAxxent Then
    chkRestore(13).Value = False
End If
chkRestore(14).Value = False
chkRestore(15).Value = False 'George 01/24/2006
chkRestore(17).Value = False    'Ticket #20536

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Data1.Refresh
DoEvents
Call Employee_Master_Integration(glbTERM_ID, EID&)

If glbWFC Then 'Ticket #16395
    'Call WFCPensionMaster(EID&, "Y")
End If

'Ticket #20270 Franks 05/05/2011
Call EEO_Process(EID&)

Call UptEmpHisTable(EID&)   'Ticket #23875 Franks 06/14/2013

If glbWFC Then
    Call locBeneGroupUpdate(EID&) 'Ticket #23247 Franks 09/17/2013
    Call mod_Upd_Pos_Budget_WFC("", "", EID&) 'Ticket #25911 Franks 12/18/2014
    Call UptMissTroyNetworkLogin(EID&) 'Ticket #28772 Franks 06/22/2016
End If

MDIMain.panHelp(0).FloodPercent = 0
Screen.MousePointer = DEFAULT
DoEvents

lblEENum = 0
lblEEName = "Employee was Reinstated"

lblEENum = 0
lblEEName = "Employee was Reinstated"
glbTERM_ID = 0
glbTERM_Seq = 0
dlpDOH.Text = ""
clpDiv = ""
txtEmpID = ""
fraLinamar.Visible = False
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).Caption = "Employee was Reinstated"
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
'cmdOK.Enabled = False

If glbWFC Then 'Ticket #24184 Franks 09/25/2013
    If glbCandidate > 0 And xHRSoftUpt Then
        Call locHRSoftAction
    End If
    
    'Ticket #24695 Franks 11/28/2013
    If GetCountryFromDiv(clpDiv1.Text) = "CANADA" Then
        MsgBox "The benefit group and Manulife fields do not get restored during the rehire process. Please go into the Status/Dates screen and re-enter this information."
    End If
End If

Me.cmdClose_Click
End Sub

Private Sub UptEmpHisTable(xEmpNo) 'Ticket #23875 Franks 06/14/2013
Dim xEmpHisDate
If Len(clpDept.Text) > 0 Then
    If Not clpDept.Text = OldDept Then
        xEmpHisDate = dlpDOH.Text
        If IsDate(dlpDeptEDate.Text) Then xEmpHisDate = dlpDeptEDate.Text
        If Not EmpHisCalc(2, xEmpNo, clpDept.Text, "", "", "", "", "", "", xEmpHisDate, , , , , , , OldDept) Then MsgBox "EMPHIS Error "
    End If
End If
If Len(clpDiv1.Text) > 0 Then
    If Not clpDiv1.Text = OldDiv Then
        xEmpHisDate = dlpDOH.Text
        If IsDate(dlpDivEDate.Text) Then xEmpHisDate = dlpDivEDate.Text
        If Not EmpHisCalc(2, xEmpNo, "", clpDiv1.Text, "", "", "", "", "", xEmpHisDate, , , , , , , OldDiv) Then MsgBox "EMPHIS Error "
    End If
End If
If Len(clpCode(1).Text) > 0 Then
    If Not clpCode(1).Text = OldLoc Then
        xEmpHisDate = dlpDOH.Text
        If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", xEmpHisDate, "LOC", clpCode(1), , , , "N", OldLoc) Then MsgBox "EMPHIS Error "
    End If
End If
If Len(clpCode(3).Text) > 0 Then
    If Not clpCode(3).Text = OldAdmin Then
        xEmpHisDate = dlpDOH.Text
        If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", xEmpHisDate, "ADMINBY", clpCode(3), , , , "N", OldAdmin) Then MsgBox "EMPHIS Error "
    End If
End If
If Len(clpCode(2).Text) > 0 Then
    If Not clpCode(2).Text = OldRegion Then
        xEmpHisDate = dlpDOH.Text
        If glbLinamar Then
            If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", xEmpHisDate, "REGION", getProductLineCodeforLinamar(clpCode(2).TransDiv & clpCode(2).Text), , , , "N") Then MsgBox "EMPHIS Error "
        Else
            If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", xEmpHisDate, "REGION", clpCode(2), , , , "N", OldRegion) Then MsgBox "EMPHIS Error "
        End If
    End If
End If
If Len(clpCode(4).Text) > 0 Then
    If Not clpCode(4).Text = OldSection Then
        xEmpHisDate = dlpDOH.Text
        If Not EmpHisCalc(2, xEmpNo, "", "", "", "", "", "", "", xEmpHisDate, "SECTION", clpCode(4), , , , "N", OldSection) Then MsgBox "EMPHIS Error "
    End If
End If
If Len(clpCode(5).Text) > 0 Then 'Union
    If Not clpCode(5).Text = OldUnion Then
        xEmpHisDate = dlpDOH.Text
        If Not EmpHisCalc(2, xEmpNo, "", "", "", "", clpCode(5).Text, "", "", xEmpHisDate, , , , , , "N", OldUnion) Then MsgBox "EMPHIS Error"
    End If
End If
If Len(clpCode(0).Text) > 0 Then 'Status
    If Not clpCode(0).Text = OldEmpStatus Then
        xEmpHisDate = dlpDOH.Text
        If Not EmpHisCalc(2, xEmpNo, "", "", clpCode(0).Text, "", "", "", "", xEmpHisDate, , , xEmpHisDate, , , "N", OldEmpStatus) Then MsgBox "EMPHIS Error"
    End If
End If

End Sub
Function getProductLineCodeforLinamar(xOrgCode) 'Ticket #23875 Franks 06/14/2013
    Dim rsTABL As New ADODB.Recordset
    Dim xNewCode
    xNewCode = xOrgCode
    rsTABL.Open "SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDRG' AND TB_KEY='" & xOrgCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If rsTABL.EOF Or rsTABL.BOF Then
        xNewCode = "ALL" & Mid(xOrgCode, 4)
    End If
    getProductLineCodeforLinamar = xNewCode
End Function

Private Sub clpDIV1_Change()
    If glbWFC Then 'Ticket #24652 Franks 12/02/2013
        Call WFC_Disp_NGSStartDate(clpDiv1.Text)
    End If
End Sub

Private Sub clpPT_Change()
    If glbWFC Then
        Call WFC_Disp_NGSStartDate(clpDiv1.Text)
    End If
End Sub

Private Sub cmdRestoreAll_Click() 'Ticket #24652 Franks 11/29/2013
Dim I As Integer
For I = 1 To 17
    If I = 5 Or I = 13 Then
    Else
        chkRestore(I).Value = True
    End If
Next
End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(Me.ActiveControl) '19Aug99 js
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRBASIC", "SELECT")
Call RollBack '29July99 js

End Sub


Private Sub Data3_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA3.error", "HRBASIC", "SELECT")
Call RollBack '29July99 js

End Sub

Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT ED_EMPNBR, ED_DOH, ED_SENDTE,ED_SIN,ED_SSN,ED_PROV,ED_COUNTRY,ED_WORKCOUNTRY,ED_DIV,ED_REGION,ED_SECTION,ED_LOC,ED_ADMINBY,ED_HIRECODE,ED_EMPTYPE,ED_TD1DOL,ED_PROVAMT,TERM_SEQ "
If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then   '
    SQLQ = SQLQ & ",ED_PAYROLL_ID,ED_BENEFIT_GROUP  "
ElseIf glbCompSerial = "S/N - 2363W" Then   'Ticket #19050 - City of Kawartha Lakes
    SQLQ = SQLQ & ",ED_FDAY "
End If
SQLQ = SQLQ & ",ED_ORG,ED_DEPTNO,ED_EMP "
SQLQ = SQLQ & " FROM Term_HREMP WHERE Term_HREMP.TERM_SEQ = " & glbTERM_Seq

If rsDATA2.State <> 0 Then rsDATA2.Close
rsDATA2.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
If Not rsDATA2.EOF Then
    If Not IsNull(rsDATA2("ED_DOH")) Then
        dlpDOH = rsDATA2("ED_DOH")
        Call setEffDates(rsDATA2("ED_DOH")) 'Ticket #23837 Franks 05/28/2013
    End If
    'Ticket #19050 - City of Kawartha Lakes
    If glbCompSerial = "S/N - 2363W" Then
        dlpLTHire = rsDATA2("ED_FDAY")
    End If
End If
If glbWFC Then 'Ticket #19231
    oDOT = getDOT(glbTERM_Seq)
    If IsNull(rsDATA2("ED_BENEFIT_GROUP")) Then SaveBGroup = "" Else SaveBGroup = rsDATA2("ED_BENEFIT_GROUP")
End If
oDOH = dlpDOH
oFday = dlpDOH

'Ticket #24317 Franks 09/17/2013 - begin
If Not IsNull(rsDATA2("ED_DEPTNO")) Then OldDept = rsDATA2("ED_DEPTNO") Else OldDept = ""
If Not IsNull(rsDATA2("ED_DIV")) Then OldDiv = rsDATA2("ED_DIV") Else OldDiv = ""
If Not IsNull(rsDATA2("ED_LOC")) Then OldLoc = rsDATA2("ED_LOC") Else OldLoc = ""
If Not IsNull(rsDATA2("ED_ADMINBY")) Then OldAdmin = rsDATA2("ED_ADMINBY") Else OldAdmin = ""
If Not IsNull(rsDATA2("ED_REGION")) Then OldRegion = rsDATA2("ED_REGION") Else OldRegion = ""
If Not IsNull(rsDATA2("ED_SECTION")) Then OldSection = rsDATA2("ED_SECTION") Else OldSection = ""
If Not IsNull(rsDATA2("ED_EMP")) Then OldEmpStatus = rsDATA2("ED_EMP") Else OldEmpStatus = ""
If Not IsNull(rsDATA2("ED_ORG")) Then OldUnion = rsDATA2("ED_ORG") Else OldUnion = ""
'Ticket #24317 Franks 09/17/2013 - end

If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then   '
    If Not IsNull(rsDATA2("ED_PAYROLL_ID")) Then
        oPayrollID = rsDATA2("ED_PAYROLL_ID")
        txtPayrollID = oPayrollID
    Else
        oPayrollID = ""
    End If
    If glbWFC Then
        locPlantCode = ""
        If Not IsNull(rsDATA2("ED_SECTION")) Then
            locPlantCode = rsDATA2("ED_SECTION")
        End If
        'Ticket #16395
        If Not IsNull(rsDATA2("ED_EMPTYPE")) Then
            If rsDATA2("ED_EMPTYPE") = "Y" Then
                comEmpType.ListIndex = 0
            End If
            If rsDATA2("ED_EMPTYPE") = "N" Then
                comEmpType.ListIndex = 1
            End If
        End If
    End If
End If
If Not rsDATA2.EOF Then
    If IsNull(rsDATA2("ED_SIN")) Then medSIN = "" Else medSIN = rsDATA2("ED_SIN")
Else
    medSIN = ""
End If

'Hemu - 06/20/2003 Begin - Ticket # 4349 To allow inputting of new employee # during rehire
If Not rsDATA2.EOF Then
    txtEmpNo.Text = ShowEmpnbr(rsDATA2("ED_EMPNBR"))
End If
'Hemu - 06/20/2003 End

If glbLinamar Then
    ' danielk - 12/31/2002 - changed from Data2 to rsDATA2
    If Not rsDATA2.EOF Then
        Call SetCountry(rsDATA2("ED_DIV"))
    End If
End If
If Not rsDATA2.EOF Then
    oSENDTE = rsDATA2!ED_SENDTE
End If
Data3.RecordSource = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & glbTERM_ID
Data3.Refresh
 
 ' out or left join query not updateable - so do straight.
EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REHIRE", "Term_HRTRMEMP", "SELECT")

Resume Next

Exit Function

End Function

Private Function EmpNoExist(xEMP)
Dim SQLQ

EmpNoExist = False
SQLQ = "SELECT ED_EMPNBR FROM HREMP"
SQLQ = SQLQ & " where ED_EMPNBR = " & xEMP & ""
Data1.RecordSource = SQLQ
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    EmpNoExist = True
    If glbLinamar Then
        If EmpNoInTerm(xEMP) Then
            EmpNoExist = False
        End If
    End If
End If

End Function

Private Sub dlpDOH_LostFocus()
Call setEffDates(dlpDOH.Text)  'Ticket #23837 Franks 05/28/2013
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMEREHIRE"
    
fglbNew = False
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEREHIRE"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
Dim Response

glbOnTop = "FRMEREHIRE"

'frmTERMEMPL.Show 1

'If glbTERM_ID = 0 Then
'    Unload Me
'    Exit Sub
'End If
Call ComEType
lblEEType.Caption = lStr("Employment Type")

Data1.ConnectionString = glbAdoIHRDB
Data3.ConnectionString = glbAdoIHRDB

Screen.MousePointer = HOURGLASS
If glbTERM_Seq = 0 Then
    frmTERMEMPL.Show 1
    If glbTERM_Seq = 0 Then
        Unload Me
        Exit Sub
    End If
Else
    'Ticket #22682 - Release 8.0 - Check if this is the newest termination record of the employee. If not then
    'warn the user.
    'Latest Termination record?
    If Not Latest_Termination(glbTERM_ID, glbTERM_Seq, glbTermDate) Then
        Response = MsgBox("This is not the most recent Termination record of the employee for Rehire." & vbCrLf & vbCrLf & "Do you wish to proceed Rehiring with this Termination Record?", vbQuestion + vbYesNo, "Confirm")
        If Response = IDNO Then
            Screen.MousePointer = DEFAULT
            
            frmTERMEMPL.Show 1
            If glbTERM_Seq = 0 Or glbTermCancel = True Then
                Unload Me
                Exit Sub
            End If
            
            'Unload Me
            'Exit Sub
        End If
    End If
End If
If EERetrieve() = False Then
    Exit Sub
Else
    'frmTERMEMPL.Show 1
    'If glbTERM_Seq = 0 Then
    '    Unload Me
    '    Exit Sub
    'End If
End If

If Len(glbTerm_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Rehire - " & Left$(glbTerm_SName, 5)
    Me.lblEEName = RTrim$(glbTerm_SName) & ", " & RTrim$(glbTerm_FName)
End If

lblEENum.Caption = glbTERM_ID

If glbDIVCount > 1 And glbLinamar Then
    clpDiv = Left(lblEEID, 3)
    txtEmpID = Mid(lblEEID, 5)
    lblEmpno.Visible = True
    frmlinamar.Visible = True
End If
If Not EmpNoExist(glbTERM_ID) Then
    If glbLinamar Then
        lblEmpno.Visible = True
        frmlinamar.Visible = True
        
        lblEmpExist.Left = 6480
    Else
        lblEmpno.Visible = True
        txtEmpNo.Visible = True
        
        'Hemu - 06/20/2003 Begin - Ticket # 4349 To allow inputting of new employee # during rehire
        txtEmpNo.Text = ""
        'Hemu - 06/20/2003 End
        
    End If
    lblEmpExist.Caption = " Employee # " & ShowEmpnbr(glbTERM_ID) & " already active - A NEW Number is required"
    lblEmpExist.Visible = True
End If

MDIMain.panHelp(0).Caption = "Complete the screen"    'laura jan 05, 1998
MDIMain.panHelp(1).Caption = " "

If Not gSec_Upd_Rehire Then
    chkRestore(0).Enabled = False
    chkRestore(1).Enabled = False
    chkRestore(2).Enabled = False
    chkRestore(3).Enabled = False
    chkRestore(4).Enabled = False
    chkRestore(5).Enabled = False
    chkRestore(6).Enabled = False
    chkRestore(7).Enabled = False 'FRANK 4/7/2000
    chkRestore(8).Enabled = False 'FRANK 4/7/2000
    chkRestore(9).Enabled = False 'FRANK 4/7/2000
    chkRestore(10).Enabled = False 'FRANK 4/7/2000
    chkRestore(11).Enabled = False 'FRANK 4/7/2000
    chkRestore(12).Enabled = False 'FRANK 4/7/2000
    chkRestore(13).Enabled = False
    chkRestore(14).Enabled = False
    chkRestore(15).Enabled = False 'George 01/24/2006
    chkRestore(16).Enabled = False 'Hemu 8/28/2010
    chkRestore(17).Enabled = False 'Ticket #20536
    dlpDOH.Enabled = False
    dlpLTHire.Enabled = False
    clpPayType.Enabled = False
    txtEmpNo.Enabled = False
    frmlinamar.Enabled = False
End If
If Not glbAxxent Then
    chkRestore(13).Value = False
    chkRestore(13).Visible = False
End If
If Not gsAttachment_DB Then  'George 01/24/2006
    chkRestore(15).Value = False
    chkRestore(15).Visible = False
End If
fraLinamar.Visible = glbLinamar
frmVadim.Visible = glbVadim

'Ticket #13165, Original Date of Hire to use the label master
lblTitle(0).Caption = "Enter New " & """" & lStr("Original Hire Date") & """"
lblTitle(6).Caption = lStr("Last Hire") 'Ticket #18668

'Hemu - Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" Then
    lblTitle(6).Visible = True
    dlpLTHire.Visible = True
    dlpLTHire.Text = dlpDOH.Text
    dlpDOH.Text = ""
    
    'Ticket #29984 - They want to change from Last Hire to Union Date
    'lblTitle(6).Caption = "Last Hire"
    dlpLTHire.DataField = "ED_UNION"
    lblTitle(6).Caption = "Union Date"
    
    Call setCaption(lblTitle(6))
    lblTitle(6).Caption = "Enter New '" & lblTitle(6).Caption & "'"
    
    lblTitle(7).Visible = True
    txtPayrollID.Visible = True
    lblTitle(7).Top = 5020
    txtPayrollID.Top = 5020
    lblPayIDExist.Top = 5020
End If
'Hemu - End

'If glbWFC Or glbCompSerial = "S/N - 2370W" Then  ' Or glbWFC
If glbCompSerial = "S/N - 2370W" Then
    lblTitle(7).Visible = True
    txtPayrollID.Visible = True
    lblTitle(7).Top = 5010 + 490
    txtPayrollID.Top = 5020 + 490
    lblPayIDExist.Top = 5020 + 490
    'Ticket #16395 - begin
    lblEEType.Top = 5360 + 490
    comEmpType.Top = 5360 + 490
    lblEEType.Visible = True '
    comEmpType.Visible = True
    'Ticket #16395 - end
End If

If glbWFC Then 'Ticket #23820 Franks 05/29/2013
    Call WFCMainScreen
    Call WFCHRSoftDispValues 'Ticket #24184 Franks 09/25/2013
End If

'Ticket #16395 - WFC Pension Outstanding Tasks By Dec2109.doc
'On Rehire (WFC only), only default checked items are Skills and Formal Education.
If glbWFC Then
    For X% = 0 To 15
        'If X% = 9 Or X% = 10 Then
        If X% = 9 Or X% = 10 Or X% = 0 Then 'Ticket #24582 added Position/Salary as defalut
            chkRestore(X%).Value = True
        Else
            chkRestore(X%).Value = False
        End If
    Next
    chkRestore(17).Value = False    'Ticket #20536 - The Hourly Entitlement was shared with Dollar Ent.
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire
    lblTitle(6).Visible = True
    lblTitle(6).FontBold = True
    dlpLTHire.Visible = True
    lblTitle(6).Caption = "Enter New " & """" & lStr("Last Hire") & """"
End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #15310
    lblTitle(6).Visible = True
    dlpLTHire.Visible = True
    dlpLTHire.Text = dlpDOH.Text
    lblTitle(6).Caption = "Last Hire"
    Call setCaption(lblTitle(6))
End If

'Jaddy begin
If glbVadim Then
    lblTitle(6).Visible = True
    dlpLTHire.Visible = True
    
    'Ticket #19050 - Not City of Kawartha Lakes
    If glbCompSerial <> "S/N - 2363W" Then
        dlpLTHire.Text = dlpDOH.Text
    End If
    
    dlpLTHire.DataField = "ED_FDAY"
    lblTitle(6).Caption = "First Day"
    Call setCaption(lblTitle(6))
    lblTitle(6).Caption = "Enter New '" & lblTitle(6).Caption & "'"
    If Vadim_PayType_field <> "" Then
        frmVadim.Visible = True
        If Not rsDATA2.EOF Then
            clpPayType.Text = rsDATA2(Vadim_PayType_field) & ""
            clpPayType.TablName = Vadim_PayType_TABLName
        End If
    End If
End If
'jaddy end

'Ticket #19310 - Samuel, Son & Co., Limited
'If glbCompSerial = "S/N - 2382W" Then
    fraBasicInfo.Visible = True
    fraBasicInfo.Left = 0 ' 240
    If glbLinamar Or glbCompSerial = "S/N - 2370W" Then
        fraBasicInfo.Top = 6580
    ElseIf glbWFC Then
        fraBasicInfo.Top = 5800 'Ticket #23820 Franks 05/29/2013
    Else
        fraBasicInfo.Top = 5850
    End If
    Call setCaption(lblTitle(11))
    Call setCaption(lblTitle(12))
    Call setCaption(lblTitle(13))
    Call setCaption(lblTitle(23))
    Call setCaption(lblTitle(24))
    Call setCaption(lblTitle(25))
    Call setCaption(lblTitle(26))
    Call setCaption(lblDeptStart)
    Call setCaption(lblDivStart)
'End If

'Ticket #19310 - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    lblTitle(11).FontBold = True
    'Ticket #20514 Franks 06/23/2011 - GL isn’t mandatory and shouldn’t be bolded
    'lblTitle(12).FontBold = True
    lblTitle(13).FontBold = True
    lblTitle(23).FontBold = True
    lblTitle(24).FontBold = True
    lblTitle(25).FontBold = True
    lblTitle(26).FontBold = True
    lblDeptStart.FontBold = True
    lblDivStart.FontBold = True
End If

If glbSamuel Then 'Ticket #21791 Franks 04/09/2012
    Call SamuelDefaultDat(glbTERM_ID, glbTERM_Seq)
End If
        
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

xUpdateable = True

Call INI_Controls(Me)

If glbLinamar Then
    If isTOUT Then
        xUpdateable = False
        MsgBox "The Transferred out employee is not allowed to rehire."
        Exit Sub
    End If
End If

End Sub

Private Function modReinMove(EID&, EESEQ&, TermDate$)
Dim X%, DtTm   As Variant, TRDesc$

Screen.MousePointer = HOURGLASS
modReinMove = False
DtTm = Now

MDIMain.panHelp(0).FloodPercent = 5

'Hemu - 06/19/2003 Begin - Since the Original Date of Hire for the Rehired employee was not changing
'                          to the new Hire Date, Ticket # 4316
'Hemu - 07/21/04 Begin - County of Essex - Modifications  - Ticket # 6549
If glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2380W" Then
    gdbAdoIhr001X.BeginTrans
    'Ticket #29984 - County of Essex - They want to change from Last Hire to Union Date
    If glbCompSerial = "S/N - 2192W" Then
        gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & ", ED_SENDTE = " & Date_SQL(dlpDOH.Text) & ", ED_UNION = " & Date_SQL(dlpLTHire) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    Else
        gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & ", ED_SENDTE = " & Date_SQL(dlpDOH.Text) & ", ED_LTHIRE = " & Date_SQL(dlpLTHire) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    End If
    gdbAdoIhr001X.CommitTrans

'Hemu - 07/21/04 End
ElseIf glbWFC Then
    gdbAdoIhr001X.BeginTrans
    'Ticket #20947 Franks 09/14/2011, MZ wants to kee the old DOH in Term_HREMP
    'gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    'Ticket #13448' Rehire will delete the Manulife Certificate#, let user to reset it
    gdbAdoIhr001X.Execute "UPDATE Term_HREMP SET ED_USER_TEXT1 = NULL,ED_USER_TEXT2=NULL,ED_USER_NUM1=NULL, ED_BENEFIT_GROUP = NULL WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    'Ticket #16392 on Rehire, make the Last Hire Date equal to the new Hire Date
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_LTHIRE = " & Date_SQL(dlpDOH.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    gdbAdoIhr001X.CommitTrans
ElseIf glbCompSerial = "S/N - 2370W" Then
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_UNION = ED_LTHIRE , ED_LTHIRE = ED_DOH, Term_HREMP.ED_EMP ='A' WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & " )"
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    gdbAdoIhr001X.CommitTrans
ElseIf glbCompSerial = "S/N - 2385W" Then 'Ticket #13165 Conservation Halton
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_UNION = " & Date_SQL(dlpDOH) & ", Term_HREMP.ED_EMP ='A' WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & " AND Term_HREMP.ED_PT ='FT')"
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_WCBCODE = '1' WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & " AND (NOT (Term_HREMP.ED_DIV ='BOAR') OR Term_HREMP.ED_DIV  IS NULL))"
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_WCBCODE = '0' WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & " AND (Term_HREMP.ED_DIV ='BOAR'))"
    gdbAdoIhr001X.CommitTrans
ElseIf glbCompSerial = "S/N - 2382W" Then
    gdbAdoIhr001X.BeginTrans
    'Ticket #18090 Samuel on Rehire, make the Last Hire Date equal to the new Hire Date
    gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_LTHIRE = " & Date_SQL(dlpDOH.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
    gdbAdoIhr001X.CommitTrans
ElseIf glbVadim Then
    'Ticket #29007 - City of Campbell River do not want auto update the Seniority Date. Also not transferring the Seniority Date to Vadim on Rehire
    If glbCompSerial = "S/N - 2458W" Then
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & ", ED_FDAY  = " & Date_SQL(dlpLTHire.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
        gdbAdoIhr001X.CommitTrans
    Else
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & ", ED_SENDTE = " & Date_SQL(dlpDOH.Text) & ", ED_FDAY  = " & Date_SQL(dlpLTHire.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
        gdbAdoIhr001X.CommitTrans
    End If
Else
    'Jaddy removed because this not make sence
    'gdbAdoIhr001X.Execute "Update Term_HREMP SET ED_DOH = " & Date_SQL(dlpDOH.Text) & ", ED_SENDTE = " & Date_SQL(dlpDOH.Text) & " WHERE (Term_HREMP.TERM_SEQ= " & EESEQ& & ")"
End If
'Hemu  - 06/19/2003 End

'Ticket #14034 - Begin Frank 12/04/07
Call UpdBenEffectiveDate(EESEQ&, dlpDOH.Text)
'Ticket #14034 - End

If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then  'County of Essex
    X% = REIN_BASIC(EID&, EESEQ&, TermDate$, txtPayrollID)
Else
    X% = REIN_BASIC(EID&, EESEQ&, TermDate$)
End If
If Not X% Then
    Exit Function
End If

X% = Reset_BASIC(EID&)
X% = RehHREMPAudit(EID&, EESEQ&)
If IsWFCNGSEmployee Then 'Ticket #25521 Franks 06/03/2014
    '"   Do not restore dependents and benefits for NGS Employees
Else
    X% = REIN_DEPEND(EID&, EESEQ&)
End If
X% = REIN_COBRA(EID&, EESEQ&)
'Ticket #19488 Frank 11/29/10
X% = REIN_HREMP_OTHER(EID&, EESEQ&)

MDIMain.panHelp(0).FloodPercent = 10

If chkRestore(1).Value = True Then
    X% = REIN_ATTENDANCE(EID&, EESEQ&)
End If

If IsWFCNGSEmployee Then 'Ticket #25521 Franks 06/03/2014
    '"   Do not restore dependents and benefits for NGS Employees
Else
    If chkRestore(2).Value = True Then
        X% = REIN_BENEFITS(EID&, EESEQ&)
        X% = RehBENEFITSAudit(EID&, EESEQ&)   'Laura jan 13, 1998
    End If
End If

MDIMain.panHelp(0).FloodPercent = 20

If chkRestore(4).Value = True Then
    X% = REIN_HealthCost(EID&, EESEQ&)
End If
MDIMain.panHelp(0).FloodPercent = 25

If chkRestore(3).Value = True Then
    X% = REIN_HealthSafety(EID&, EESEQ&)
    X% = REIN_OHS_CONTACT(EID&, EESEQ&)
    X% = REIN_OHS_CORRECTIVE(EID&, EESEQ&)
    X% = REIN_OHS_ROOT_CAUSES(EID&, EESEQ&)
    X% = REIN_OHS_CLAIM_MEDICAL(EID&, EESEQ&)
    X% = REIN_OHS_FORM7_SECTIONS(EID&, EESEQ&)
    X% = REIN_OHS_FORM9(EID&, EESEQ&)
End If
MDIMain.panHelp(0).FloodPercent = 30

If chkRestore(0).Value = True Then
    X% = REIN_JOB(EID&, EESEQ&)
    X% = RehJOBAudit(EID&, EESEQ&)     'laura jan 13, 1998
    MDIMain.panHelp(0).FloodPercent = 40
    
    X% = REIN_PERFORM(EID&, EESEQ&)
    MDIMain.panHelp(0).FloodPercent = 60
    
    X% = REIN_SALARY(EID&, EESEQ&)
    X% = RehSALARYAudit(EID&, EESEQ&)   'laura jan 13, 1998
    MDIMain.panHelp(0).FloodPercent = 75
End If

If chkRestore(6).Value = True Then  'laura nov 5, 1997
    X% = REIN_EDUCSEM(EID&, EESEQ&)
End If

If chkRestore(7).Value = True Then  'FRANK 4/7/2000
    X% = REIN_COMMENTS(EID&, EESEQ&)
End If

If chkRestore(8).Value = True Then  'FRANK 4/7/2000
    X% = REIN_EARN(EID&, EESEQ&)
End If

If chkRestore(9).Value = True Then  'FRANK 4/7/2000
    X% = REIN_EDU(EID&, EESEQ&)
End If

If chkRestore(10).Value = True Then  'FRANK 4/7/2000
    X% = REIN_EMPSKL(EID&, EESEQ&)
End If

If chkRestore(11).Value = True Then  'FRANK 4/7/2000
    X% = REIN_TRADE(EID&, EESEQ&)
End If

If chkRestore(12).Value = True Then  'FRANK 4/7/2000
    X% = REIN_DOLENT(EID&, EESEQ&)
    
    'Ticket #28789 - Actual Amounts Details
    X% = REIN_DOLENT_ACTDTL(EID&, EESEQ&)
End If

If chkRestore(17).Value = True Then  'Ticket #20536
    X% = REIN_ENTHRS(EID&, EESEQ&)
End If

If glbAxxent Then 'Ticket #25023 Franks 01/30/2014
    If chkRestore(13).Value = True Then
        X% = REIN_RSP(EID&, EESEQ&)
    End If
End If

If chkRestore(14).Value = True Then
    X% = REIN_COUNSEL(EID&, EESEQ&)
    
    'Ticket #23409 - Samuel - Add Discipline Audit
    If glbSamuel Then
        X% = RehCOUNSELAudit(EID&, EESEQ&)
    End If
End If

X% = REIN_USERDEFINED(EID&, EESEQ&)     'Hemu - User Defined Table

X% = REIN_Profit_Sharing(EID&, EESEQ&)     'Franks - 07/25/2011 - Ticket #20052

X% = REIN_EMP_FLAGS(EID&, EESEQ&)     'Hemu - Ticket #24065

X% = REIN_HREEO(EID&, EESEQ&) 'Ticket #25669 Franks 06/24/2014

If chkRestore(15).Value = True Then 'George 01/24/2006
    If gsAttachment_DB Then
        X% = REIN_HRDOC_EMP(EID&, EESEQ&)
        X% = REIN_HRDOC_JOB_HISTORY(EID&, EESEQ&)
        X% = REIN_HRDOC_COMMENTS(EID&, EESEQ&)
        X% = REIN_HRDOC_HEALTH_SAFETY(EID&, EESEQ&)
        X% = REIN_HRDOC_HEALTH_SAFETY_2(EID&, EESEQ&)
        If glbWSIBModule Then
            X% = REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7(EID&, EESEQ&)
            X% = REIN_HRDOC_OHS_WRITTEN_OFFER(EID&, EESEQ&)
        End If
        X% = REIN_HRDOC_COUNSEL(EID&, EESEQ&)
        X% = REIN_HRDOC_PERFORM_HISTORY(EID&, EESEQ&)
        'add missed tables by Frank 01/21/10 Ticket #17894 - begin
        X% = REIN_HRDOC_EDSEM(EID&, EESEQ&)
        X% = REIN_HRDOC_EDSEM_RETEST(EID&, EESEQ&)
        X% = REIN_HRDOC_HREDU(EID&, EESEQ&)
        X% = REIN_HRDOC_DOLENT(EID&, EESEQ&)
        X% = REIN_HRDOC_TRADE(EID&, EESEQ&)
        X% = REIN_HRDOC_ATTENDANCE(EID&, EESEQ&)
        X% = REIN_HRDOC_EMP_FLAGS(EID&, EESEQ&)
        'add missed tables by Frank 01/21/10 Ticket #17894 - end
        
        'Release 8.1
        X% = REIN_HRDOC_HREMP_OTHER(EID&, EESEQ&)
    End If
End If

If chkRestore(16).Value = True Then
    X% = REIN_FOLLOW_UP(EID&, EESEQ&)
End If

X% = REIN_SUCCESSION(EID&, EESEQ&) 'George Apr 4,2006 #10595
X% = REIN_LANGUAGE(EID&, EESEQ&) 'George Apr 4,2006 #10595
X% = REIN_HREMPHIS(EID&, EESEQ&)

If glbLinamar Then
    X% = REIN_LN_EMPSKL(EID&, EESEQ&)

    'Ticket #13799 - Restore the Employee Photo - Change the employee # to new #.
    X% = REIN_HR_PHOTO(EID&, EESEQ&, lblEENum)
Else
    'Ticket #20367 - Jerry said we should not delete the Photo and also should Rehire and
    'we should be able to see Photo on Demographics screen of the Terminated employee
    'Ticket #13799 - Restore the Employee Photo - Change the employee # to new #.
    X% = REIN_HR_PHOTO(EID&, EESEQ&, lblEENum)
End If

'Ticket #25459 - Terminate ESS and TS employee records as well
'Web Modules Begin
X% = REIN_VACTIMEOFF_REQ(EID&, EESEQ&)
X% = REIN_VACTIMEOFF_REQ_ARCHIVE(EID&, EESEQ&)
X% = REIN_VACTIMEOFF_REQ_WRK(EID&, EESEQ&)
X% = REIN_REQAUDIT(EID&, EESEQ&)
X% = REIN_TIMESHEET(EID&, EESEQ&)
X% = REIN_TIMESHEET_ARCHIVE(EID&, EESEQ&)
'Web Modules End

X% = RemoveHREMPEQU_DOT(EID&)

If glbLambton Then
    X% = RemoveCurrentFlag(EID&)
End If
    
Dim HRChanges As New Collection
Call isChanged_Field(HRChanges, oDOH, dlpDOH)
Call isChanged_Field(HRChanges, oFday, dlpLTHire)
Call Passing_Changes(HRChanges, Rehire, "M", Date, EID&)

Call AddNewPayrollEmp(Rehire, Date, EID&, "")

If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #24729 01/21/2014 Franks
    Call FamilDayEmpUpt(EID&)
End If

modReinMove = True

Screen.MousePointer = DEFAULT

Exit Function

modReinMoveErr_Msg:
Screen.MousePointer = DEFAULT
MsgBox "Problem Creating Audit record - Termination Aborted"

End Function

Private Function RehBENEFITSAudit(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
'Dim TIHR_DB As Database

'Laura jan 13, 1998
'update HRAUDIT from Term_BENEFITS
RehBENEFITSAudit = False

On Error GoTo RehBENEFITSAudit_Err

SQLQ = "INSERT INTO HRAUDIT ( AU_COMPNO, AU_EMPNBR, AU_BCODE, "
SQLQ = SQLQ & "AU_COVER, AU_BAMT, AU_EDATE, AU_PCE, AU_PCC, "
SQLQ = SQLQ & "AU_TCOST, AU_UNITCOST, AU_PREMIUM, AU_PER, "
SQLQ = SQLQ & "AU_PPAMT, AU_MTHCCOST, AU_MTHECOST, "
SQLQ = SQLQ & "AU_TAXBEN, "
SQLQ = SQLQ & "AU_LDATE, AU_LTIME, AU_LUSER, AU_TYPE, AU_UPLOAD ) "
SQLQ = SQLQ & "SELECT Term_HRBENFT.BF_COMPNO, "
SQLQ = SQLQ & EEID & " , Term_HRBENFT.BF_BCODE, "
SQLQ = SQLQ & "Term_HRBENFT.BF_COVER, Term_HRBENFT.BF_AMT, "
SQLQ = SQLQ & "Term_HRBENFT.BF_EDATE, Term_HRBENFT.BF_PCE, "
SQLQ = SQLQ & "Term_HRBENFT.BF_PCC, Term_HRBENFT.BF_TCOST, "
SQLQ = SQLQ & "Term_HRBENFT.BF_UNITCOST, Term_HRBENFT.BF_PREMIUM, "
SQLQ = SQLQ & "Term_HRBENFT.BF_PER, Term_HRBENFT.BF_PPAMT, "
SQLQ = SQLQ & "Term_HRBENFT.BF_MTHCCOST, Term_HRBENFT.BF_MTHECOST, "
SQLQ = SQLQ & "Term_HRBENFT.BF_TAXBEN, "
SQLQ = SQLQ & Date_SQL(Date) & " As AU_LDATE, '"

SQLQ = SQLQ & Time$ & "' As AU_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' As AU_LUSER, 'A' As AU_TYPE,'N' AS AU_UPLOAD FROM Term_HRBENFT "
SQLQ = SQLQ & "WHERE (Term_HRBENFT.TERM_SEQ =" & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

RehBENEFITSAudit = True
'TIHR_DB.Close

Exit Function

RehBENEFITSAudit_Err:
glbFrmCaption$ = "Rehire Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehBENEFITSAudit", "Term_HRBENFT", "Insert")
Call RollBack '29July99 js

End Function

Private Function RehHREMPAudit(EEID&, EESEQ&)
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim xProvNbr, xADD, xPROV
Dim Langs 'George Apr 4,2006 #10574
'Dim TIHR_DB As Database

On Error GoTo RehHREMPAudit_ERR

RehHREMPAudit = False


'Rehire using enough fields to make * worth it, Ticket#9899
rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'Ticket# 4032 For 1010 demo DB, show "SYSTEM ERROR - READING TERM_EMP"
'rsTC.Open "select * from HREMP where ED_EMPNBR = " & EEID&, gdbAdoIhr001, adOpenKeyset
'Ticket #19310 - Samuel, Son & Co., Limited
If glbLinamar Or glbCompSerial = "S/N - 2382W" Then
    rsTC.Open "select * from HREMP where ED_EMPNBR = " & EEID&, gdbAdoIhr001, adOpenKeyset
Else
    rsTC.Open "Term_HREMP", gdbAdoIhr001X, adOpenKeyset, , adCmdTableDirect
    rsTC.Find "TERM_SEQ = " & EESEQ
End If

If rsTC.EOF Then
    MsgBox "SYSTEM ERROR - READING TERM_EMP"
    Exit Function
End If

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_COMPNO") = "001"
If glbCompSerial = "S/N - 2372W" Then 'Town of Bradford West Gwillimbury
    If locNewHire Then
        rsTA("AU_NEWEMP") = "Y"
    Else
        rsTA("AU_NEWEMP") = "N"
    End If
ElseIf glbSamuel Then
    'Ticket #21791 Franks 04/09/2012
    'Insync interface treat rehire as modification, not new hire
    rsTA("AU_NEWEMP") = "N"
ElseIf glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #26922 04/13/2015 Franks
    'treat rehire as modification, not new hire
    rsTA("AU_NEWEMP") = "N"
ElseIf glbCompSerial = "S/N - 2353W" Then  'Let's Talk Science Ticket #28821 09/01/2016 Franks
    'treat rehire as modification, not new hire
    rsTA("AU_NEWEMP") = "N"
Else
    rsTA("AU_NEWEMP") = "Y"
End If
If glbWFC Then 'Ticket #16749
    'rsTA("AU_TYPE") = "H"
    'Ticket #19231 10/06/2010 Frank - begin
    'On rehire, if the date of termination is less than Jan 1, 2009 transfer a new hire record
    'to Payforce and not a job change.
    If Not IsDate(oDOT) Then
        oDOT = Date
    End If
    If CVDate(oDOT) < CVDate("01/01/2009") Then
        rsTA("AU_TYPE") = "A"
    Else
        rsTA("AU_TYPE") = "H"
    End If
    'Ticket #19231 10/06/2010 Frank - end
Else
    rsTA("AU_TYPE") = "A"
End If
If Len(clpDiv1.Text) > 0 Then 'Ticket #20635 Franks 07/15/2011
    rsTA("AU_DIV") = clpDiv1.Text
    rsTA("AU_DIVUPL") = clpDiv1.Text
Else
    rsTA("AU_DIV") = rsTC("ED_DIV")
    rsTA("AU_DIVUPL") = rsTC("ED_DIV")
End If
If Len(clpCode(1).Text) > 0 Then 'Ticket #20635 Franks 07/15/2011
    rsTA("AU_LOC") = clpCode(1).Text
Else
    rsTA("AU_LOC") = rsTC("ED_LOC")
End If
rsTA("AU_EMPNBR") = EEID 'rsTC("ED_EMPNBR") 'Ticket
rsTA("AU_TITLE") = rsTC("ED_TITLE")
rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
rsTA("AU_FNAME") = rsTC("ED_FNAME")
rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
rsTA("AU_CITY") = rsTC("ED_CITY")
rsTA("AU_PROV") = rsTC("ED_PROV")
rsTA("AU_PCODE") = rsTC("ED_PCODE")
rsTA("AU_PHONE") = rsTC("ED_PHONE")
rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
rsTA("AU_PROVRES") = rsTC("ED_PROVEMP")
rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
rsTA("AU_UIC") = rsTC("ED_UIC")
rsTA("AU_PENSION") = rsTC("ED_PENSION")
rsTA("AU_CPP") = rsTC("ED_CPP")
rsTA("AU_GROSSCD") = rsTC("ED_GROSSCD")
rsTA("AU_GARN") = rsTC("ED_GARN")
rsTA("AU_ELIGIBLE") = rsTC("ED_ELIGIBLE")
rsTA("AU_EARLYR") = rsTC("ED_EARLYR")
rsTA("AU_NORMALR") = rsTC("ED_NORMALR")
rsTA("AU_LATESTR") = rsTC("ED_LATESTR")
rsTA("AU_SIN") = rsTC("ED_SIN")
'rsTA("AU_PT") = rsTC("ED_PT")
rsTA("AU_PTUPL") = rsTC("ED_PT")
rsTA("AU_EMPTYPE") = rsTC("ED_EMPTYPE")
rsTA("AU_SEX") = rsTC("ED_SEX")
If rsTC("ED_SMOKER") <> 0 Then
    rsTA("AU_SMOKER") = "Yes"
Else
    rsTA("AU_SMOKER") = "No"
End If
rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
If Len(clpDept.Text) > 0 Then 'Ticket #20635 Franks 07/15/2011
    rsTA("AU_DEPTNO") = clpDept.Text
Else
    rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
End If
rsTA("AU_DOB") = rsTC("ED_DOB")
If IsDate(dlpDOH.Text) Then
    rsTA("AU_DOH") = dlpDOH.Text
Else
    rsTA("AU_DOH") = rsTC("ED_DOH")
End If
rsTA("AU_SENDTE") = rsTC("ED_SENDTE")
rsTA("AU_LTHIRE") = rsTC("ED_LTHIRE")
rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
rsTA("AU_BANK") = rsTC("ED_BANK")
rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
rsTA("AU_BANK2") = rsTC("ED_BANK2")
rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
rsTA("AU_PCDEPOSIT2") = rsTC("ED_PCDEPOSIT2")
rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
rsTA("AU_BANK3") = rsTC("ED_BANK3")
rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
rsTA("AU_SUPCODE") = rsTC("ED_SUPCODE")
rsTA("AU_DDI") = rsTC("ED_DDI")
rsTA("AU_WCB") = rsTC("ED_WCB")
rsTA("AU_UNION") = rsTC("ED_UNION")
rsTA("AU_TD1") = rsTC("ED_TD1")
rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
rsTA("AU_TD3") = rsTC("ED_TD3")
rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
rsTA("AU_VACPC") = rsTC("ED_VACPC")
rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
rsTA("AU_FDAY") = rsTC("ED_FDAY")
rsTA("AU_LDAY") = rsTC("ED_LDAY")
rsTA("AU_OMDAY") = rsTC("ED_OMERS")

rsTA("AU_INTEL") = rsTC("ED_INTEL")
'George Apr 4,2006 #10574
'rsTA("AU_LANG1") = rsTC("ED_LANG1")
'rsTA("AU_LANG2") = rsTC("ED_LANG2")
Langs = Split(getLanguage(rsTC("ED_EMPNBR")), "|")
If Langs(0) <> "NoLang1" Then rsTA("AU_LANG1") = Langs(0) '0 is for ED_Lang1
If Langs(1) <> "NoLang2" Then rsTA("AU_LANG2") = Langs(1) '1 is for ED_Lang2
'George Apr 4,2006 #10574
rsTA("AU_EMAIL") = rsTC("ED_EMAIL")
rsTA("AU_WITHSPOUSE") = rsTC("ED_WITHSPOUSE")
rsTA("AU_EXPYEAR") = rsTC("ED_EXPYEAR")
If Len(clpCode(3).Text) > 0 Then 'Ticket #20635 Franks 07/15/2011
    rsTA("AU_ADMINBY") = clpCode(3).Text
Else
    rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
End If
rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
rsTA("AU_SSN") = rsTC("ED_SSN")
If Len(clpGLNum.Text) > 0 Then 'Ticket #20635 Franks 07/15/2011
    rsTA("AU_DEPT_GL") = clpGLNum.Text
Else
    rsTA("AU_DEPT_GL") = rsTC("ED_GLNO")
End If
If Len(clpCode(2).Text) > 0 Then 'Ticket #20635 Franks 07/15/2011
    rsTA("AU_REGION") = clpCode(2).Text
Else
    rsTA("AU_REGION") = rsTC("ED_REGION")
End If
If Len(clpCode(4).Text) > 0 Then
    rsTA("AU_SECTION") = clpCode(4).Text
Else
    rsTA("AU_SECTION") = rsTC("ED_SECTION")
End If
'Ticket #23820 Franks 05/29/2013 - begin
'rsTA("AU_EMP") = rsTC("ED_EMP")
If Len(clpCode(0).Text) > 0 Then rsTA("AU_EMP") = clpCode(0).Text Else rsTA("AU_EMP") = rsTC("ED_EMP")
'rsTA("AU_ORG") = rsTC("ED_ORG")
If Len(clpCode(5).Text) > 0 Then rsTA("AU_ORG") = clpCode(5).Text Else rsTA("AU_ORG") = rsTC("ED_ORG")
'rsTA("AU_PT") = rsTC("ED_PT")
If Len(clpPT.Text) > 0 Then rsTA("AU_PT") = clpPT.Text Else rsTA("AU_PT") = rsTC("ED_PT") 'Ticket #25562 Franks 06/17/2014
'Ticket #23820 Franks 05/29/2013 - end
If glbWFC Then 'Ticket #23247 Franks 09/17/2013
    If Len(NewPayGroup) > 0 Then rsTA("AU_VADIM2") = NewPayGroup
    If Len(NewNGSSub) > 0 Then rsTA("AU_VADIM1") = NewNGSSub
End If
rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")

If glbLinamar Then
    rsTA("AU_HOMELINE") = rsTC("ED_HOMELINE")
    rsTA("AU_HOMESHIFT") = rsTC("ED_HOMESHIFT")
    rsTA("AU_HOMEOPRTNBR") = rsTC("ED_HOMEOPRTNBR")
    rsTA("AU_HOMEWRKCNT") = rsTC("ED_HOMEWRKCNT")
    rsTA("AU_ExtrAnn") = rsTC("ED_ExtrAnn")
    rsTA("AU_QTBTORRSP") = rsTC("ED_QTBTORRSP")
    rsTA("AU_LOCKER") = rsTC("ED_LOCKER")
    rsTA("AU_COMBINATION") = rsTC("ED_COMBINATION")
End If

'If glbSoroc Or glbSyndesis Then rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then  'County of Essex
    rsTA("AU_PAYROLL_ID") = txtPayrollID
Else
    If Not IsNull(rsTC("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
End If

If IsDate(dlpDeptEDate.Text) Then  'Ticket #20635 Franks 07/15/2011
    rsTA("AU_DEPTEDATE") = dlpDeptEDate.Text
Else
    rsTA("AU_DEPTEDATE") = rsTC("ED_DEPTEDATE")
End If
If IsDate(dlpDivEDate.Text) Then  'Ticket #20635 Franks 07/15/2011
    rsTA("AU_DIVEDATE") = dlpDivEDate.Text
Else
    rsTA("AU_DIVEDATE") = rsTC("ED_DIVEDATE")
End If
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA.Update
If glbLinamar Then
    Dim xKey, xJob
    xKey = "E" & rsTC!ED_EMPNBR
    rsTB.Open "SELECT JH_JOB FROM Term_JOB_HISTORY WHERE JH_CURRENT<>0 AND Term_SEQ=" & EESEQ, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        xJob = rsTB!JH_JOB
    Else
        xJob = ""
    End If
    rsTA.Close
    rsTA.Open "LN_TRALOG", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsTA.AddNew
    
    rsTA!TL_COMPNO = "001"
    rsTA!TL_EMPNBR = rsTC!ED_EMPNBR
    rsTA!TL_SURNAME = rsTC!ED_SURNAME
    rsTA!TL_FNAME = rsTC!ED_FNAME
    rsTA!TL_DOH = rsTC!ED_DOH
    rsTA!TL_JOB = xJob
    
    rsTA!TL_TYPE = "R-HI"
    
    rsTA!TL_OLDDIV = Right(rsDATA2!ED_EMPNBR, 3)
    rsTA!TL_OLDEMPNBR = rsDATA2!ED_EMPNBR
    rsTA!TL_OLDDIVEDATE = oSENDTE

    
    rsTA!TL_NEWDIV = rsTC!ED_DIV
    rsTA!TL_NEWEMPNBR = rsTC!ED_EMPNBR
    rsTA!TL_NEWDIVEDATE = rsTC!ED_SENDTE
    
    rsTA!TL_TOREASON_TABL = "TERM"
    rsTA!TL_TIREASON_TABL = "SDJC"
    rsTA!TL_TERM_SEQ = EESEQ
    rsTA!TL_TCOMPLETE = "Y"
    
    rsTA!TL_KEY = xKey
    rsTA!TL_CURRENTDIV = rsTC!ED_DIV
    
    rsTA("TL_LDATE") = Format(Now, "SHORT DATE")
    rsTA("TL_LUSER") = glbUserID
    rsTA("TL_LTIME") = Time$
    rsTA.Update
    rsTA.Close
    rsTA.Open "SELECT TL_KEY,TL_CURRENTDIV FROM LN_TRALOG WHERE TL_KEY='T" & EESEQ & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do Until rsTA.EOF
        rsTA!TL_KEY = xKey
        rsTA!TL_CURRENTDIV = rsTC!ED_DIV
        rsTA.Update
        rsTA.MoveNext
    Loop

'    gdbAdoIhr001.Execute "UPDATE LN_TRALOG SET TL_KEY='" & xKEY & "' WHERE TL_KEY='T" & EESEQ & "'"
End If
RehHREMPAudit = True

Exit Function

RehHREMPAudit_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js

End Function

Private Function RehJOBAudit(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
'Dim TIHR_DB As Database

'Laura jan 13, 1998
'update HRAUDIT from Term_JOB_HISTORY
RehJOBAudit = False

On Error GoTo RehJOBAudit_Err

SQLQ = "INSERT INTO HRAUDIT ( AU_COMPNO, AU_EMPNBR, AU_WHRS, "
SQLQ = SQLQ & "AU_DHRS, AU_PHRS, AU_SJDATE, "
If glbLinamar Then SQLQ = SQLQ & "AU_LEADHAND,AU_LABOURCD,"
SQLQ = SQLQ & "AU_PAYROLL_ID," 'If glbSoroc Or glbSyndesis Then
SQLQ = SQLQ & "AU_JOB, AU_LDATE, AU_LTIME, AU_LUSER, AU_TYPE , AU_UPLOAD) "
SQLQ = SQLQ & "SELECT Term_JOB_HISTORY.JH_COMPNO, "
SQLQ = SQLQ & EEID & " , Term_JOB_HISTORY.JH_WHRS, "
SQLQ = SQLQ & "Term_JOB_HISTORY.JH_DHRS, Term_JOB_HISTORY.JH_PHRS, "
SQLQ = SQLQ & "Term_JOB_HISTORY.JH_SDATE, "
If glbLinamar Then SQLQ = SQLQ & "Term_JOB_HISTORY.JH_LEADHAND,Term_JOB_HISTORY.JH_LABOURCD,"
If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then  'County of Essex
    SQLQ = SQLQ & "'" & txtPayrollID & "',"
Else
    SQLQ = SQLQ & "Term_HREMP.ED_PAYROLL_ID," 'If glbSoroc Or glbSyndesis Then
End If
SQLQ = SQLQ & "Term_JOB_HISTORY.JH_JOB, "

'Ticket #23485 - Town of Orangeville
If glbCompSerial = "S/N - 2383W" Then
    SQLQ = SQLQ & "JH_SDATE As AU_LDATE, '"
Else
    SQLQ = SQLQ & Date_SQL(Date) & " As AU_LDATE, '"
End If

SQLQ = SQLQ & Time$ & "' As AU_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' As AU_LUSER, 'A' As AU_TYPE,'N' AS AU_UPLOAD FROM Term_JOB_HISTORY "
If glbOracle Then
    SQLQ = SQLQ & ",Term_HREMP WHERE Term_JOB_HISTORY.TERM_SEQ = Term_HREMP.TERM_SEQ " 'If glbSoroc Or glbSyndesis Then
    SQLQ = SQLQ & " AND (Term_JOB_HISTORY.TERM_SEQ =" & EESEQ & ") AND JH_CURRENT<>0"
Else
    SQLQ = SQLQ & "INNER JOIN Term_HREMP ON Term_JOB_HISTORY.TERM_SEQ = Term_HREMP.TERM_SEQ " 'If glbSoroc Or glbSyndesis Then
    SQLQ = SQLQ & "WHERE (Term_JOB_HISTORY.TERM_SEQ =" & EESEQ & ") AND JH_CURRENT<>0"
End If

gdbAdoIhr001X.Execute SQLQ

RehJOBAudit = True
'TIHR_DB.Close

Exit Function

RehJOBAudit_Err:
glbFrmCaption$ = "Rehire Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehJOBAudit", "Term_JOB_HISTORY", "Insert")
Call RollBack '29July99 js

End Function

Private Function RehSALARYAudit(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
'Dim TIHR_DB As Database
'Laura jan 13, 1998
'update HRAUDIT from Term_SALARY_HISTORY

RehSALARYAudit = False

On Error GoTo RehSALARYAudit_Err

SQLQ = "INSERT INTO HRAUDIT ( AU_COMPNO, AU_EMPNBR, AU_SALCD, "
SQLQ = SQLQ & "AU_PAYP, AU_SEDATE, AU_SNDATE, "
SQLQ = SQLQ & "AU_SALARY, "
SQLQ = SQLQ & "AU_PAYROLL_ID," 'If glbSoroc Or glbSyndesis Then
SQLQ = SQLQ & "AU_SJDATE, AU_JOB, AU_LDATE, AU_LTIME, AU_LUSER, AU_TYPE, AU_UPLOAD ) "
SQLQ = SQLQ & "SELECT Term_SALARY_HISTORY.SH_COMPNO, "
SQLQ = SQLQ & EEID & " , Term_SALARY_HISTORY.SH_SALCD, "
SQLQ = SQLQ & "Term_SALARY_HISTORY.SH_PAYP, Term_SALARY_HISTORY.SH_EDATE, "
SQLQ = SQLQ & "Term_SALARY_HISTORY.SH_NEXTDAT, Term_SALARY_HISTORY.SH_SALARY, "
If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then  'County of Essex
    SQLQ = SQLQ & "'" & txtPayrollID & "',"
Else
    SQLQ = SQLQ & "Term_HREMP.ED_PAYROLL_ID," 'If glbSoroc Or glbSyndesis Then
End If
SQLQ = SQLQ & "Term_SALARY_HISTORY.SH_SDATE, Term_SALARY_HISTORY.SH_JOB, "
'Ticket #20843 franks 08/23/11, do not use SH_EDATE as AU_LDATE, it caused the payroll problem
'If glbCompSerial = "S/N - 2290W" Or glbCompSerial = "S/N - 2370W" Then
'    SQLQ = SQLQ & Date_SQL(Date) & " As AU_LDATE, "
'Else
'    SQLQ = SQLQ & "SH_EDATE As AU_LDATE, "
'End If
'Ticket #23485 - Town of Orangeville
If glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2383W" Then  'Samuel  - Ticket #21104 Franks 10/25/2011
    SQLQ = SQLQ & "SH_EDATE As AU_LDATE, "
Else
    SQLQ = SQLQ & Date_SQL(Date) & " As AU_LDATE, "
End If
SQLQ = SQLQ & "'" & Time$ & "' As AU_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' As AU_LUSER, 'A' As AU_TYPE,'N' AS AU_UPLOAD FROM Term_SALARY_HISTORY "

If glbOracle Then
    SQLQ = SQLQ & " ,Term_HREMP WHERE Term_SALARY_HISTORY.TERM_SEQ = Term_HREMP.TERM_SEQ " 'If glbSoroc Or glbSyndesis Then
    SQLQ = SQLQ & " AND (Term_SALARY_HISTORY.TERM_SEQ =" & EESEQ & ") AND SH_CURRENT<>0"
Else
    SQLQ = SQLQ & "INNER JOIN Term_HREMP ON Term_SALARY_HISTORY.TERM_SEQ = Term_HREMP.TERM_SEQ " 'If glbSoroc Or glbSyndesis Then
    SQLQ = SQLQ & "WHERE (Term_SALARY_HISTORY.TERM_SEQ =" & EESEQ & ") AND SH_CURRENT<>0"
End If

gdbAdoIhr001X.Execute SQLQ
RehSALARYAudit = True
'TIHR_DB.Close

Exit Function

RehSALARYAudit_Err:
glbFrmCaption$ = "Rehire Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehSALARYAudit", "Term_SALARY_HISTORY", "Insert")
Call RollBack '29July99 js

End Function

Private Function RehCOUNSELAudit(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String

RehCOUNSELAudit = False

On Error GoTo RehCOUNSELAudit_Err

SQLQ = "INSERT INTO HRAUDIT_COUNSEL (AU_COMPNO, AU_EMPNBR, AU_COUDATE, AU_INCDATE, "
SQLQ = SQLQ & "AU_TYPE, AU_REASON, AU_FOLLOWUPD1, AU_FOLLOWUPD2, AU_FOLLOWUPD3, "
SQLQ = SQLQ & "AU_COMMENTS, AU_COUBY, AU_ATTDATE, AU_ATTREASON, AU_DATE1, AU_COMPLETED, AU_EMP_RESPONSE, "
SQLQ = SQLQ & "AU_EMP_AGREED, AU_EMP_DECLINED, AU_LDATE, AU_LTIME, AU_LUSER, AU_TRANS_TYPE, AU_UPLOAD) "
SQLQ = SQLQ & "SELECT Term_HR_COUNSEL.CL_COMPNO, "
SQLQ = SQLQ & EEID & " , Term_HR_COUNSEL.CL_COUDATE, Term_HR_COUNSEL.CL_INCDATE, Term_HR_COUNSEL.CL_TYPE, "
SQLQ = SQLQ & "Term_HR_COUNSEL.CL_REASON, Term_HR_COUNSEL.CL_FOLLOWUPD1, Term_HR_COUNSEL.CL_FOLLOWUPD2, Term_HR_COUNSEL.CL_FOLLOWUPD3, "
SQLQ = SQLQ & "Term_HR_COUNSEL.CL_COMMENTS, Term_HR_COUNSEL.CL_COUBY, Term_HR_COUNSEL.CL_ATTDATE, Term_HR_COUNSEL.CL_ATTREASON, "
SQLQ = SQLQ & "Term_HR_COUNSEL.CL_DATE1, Term_HR_COUNSEL.CL_COMPLETED, Term_HR_COUNSEL.CL_EMP_RESPONSE,"
SQLQ = SQLQ & "Term_HR_COUNSEL.CL_EMP_AGREED, Term_HR_COUNSEL.CL_EMP_DECLINED, "
SQLQ = SQLQ & Date_SQL(Date) & " As AU_LDATE, '"
SQLQ = SQLQ & Time$ & "' As AU_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' As AU_LUSER, 'R' As AU_TRANS_TYPE,'N' AS AU_UPLOAD FROM Term_HR_COUNSEL "
SQLQ = SQLQ & "WHERE (Term_HR_COUNSEL.TERM_SEQ =" & EESEQ & ") "

gdbAdoIhr001X.Execute SQLQ

RehCOUNSELAudit = True


Exit Function

RehCOUNSELAudit_Err:
glbFrmCaption$ = "Rehire Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "RehCOUNSELAudit", "Term_HR_COUNSEL", "Insert")
Call RollBack

End Function

Private Function RemoveHREMPEQU_DOT(EmpN As Long)
Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset

SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_EMPNBR = "
SQLQ = SQLQ & EmpN

dynEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If dynEmp.RecordCount > 0 Then
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    'SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = Null "
    SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = Null, EQ_TYPE = 'A' "
    SQLQ = SQLQ & "WHERE HREMPEQU.EQ_EMPNBR = " & EmpN
    gdbAdoIhr001.Execute SQLQ
End If

End Function

Private Function RemoveCurrentFlag(EmpN As Long)
Dim SQLQ As String
SQLQ = "UPDATE HR_JOB_HISTORY SET JH_CURRENT=0 WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpN
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_CURRENT=0 WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpN
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
SQLQ = "UPDATE HR_PERFORM_HISTORY SET PH_CURRENT=0 WHERE PH_CURRENT<>0 AND PH_EMPNBR=" & EmpN
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    glbOnTop = ""
    glbTERM_Seq = 0
    glbTERM_ID = 0  'Ticket #24386 - To avoid showing a blank data screen
End Sub

Private Sub Form_Resize()
Dim c As Long

On Error GoTo Err_Resize

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 10000 Then
        scrControl.Value = 0
        
        frRehire.Top = 1080
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 2500
        
        If Me.Height < 8500 Then
            scrControl.Max = 4100
        Else
            scrControl.Max = 1500
        End If
        
    End If


'    'Horizontal Scroll
'    scrHScroll.Width = Me.Width - 200
'    If Me.Width >= 11190 Then '9700 Then
'        scrHScroll.Value = 0
'        scrHScroll.Visible = False
'    Else
'        scrHScroll.Visible = True
'        scrHScroll.Top = Me.Height - 700
'        scrHScroll.Width = Me.Width - 120
'    End If
    
End If

exH:
    Exit Sub
    
Err_Resize:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Rehire", "Update")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEREHIRE = Nothing
    glbTERM_Seq = 0
    glbTERM_ID = 0  'Ticket #24386 - To avoid showing a blank data screen
End Sub

Private Sub lblEENum_Change()
lblEEID.Caption = ShowEmpnbr(lblEENum)
End Sub

Private Function RollBack()

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub scrControl_Change()
frRehire.Top = 1080 - scrControl.Value
End Sub

Private Sub txtEmpNo_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Function Reset_BASIC(EID&)
Dim rsEmp As New ADODB.Recordset
Dim rsPA As New ADODB.Recordset
Dim xTDate

Screen.MousePointer = HOURGLASS

Reset_BASIC = False

'added by Bryan 24/Apr/2006 Ticket#10313
Pause (3)

'rsEmp.Open "SELECT ED_DIV,ED_DOH,ED_SENDTE,ED_COUNTRY,ED_PROV,ED_SIN,ED_LTHIRE,ED_FDAY,ED_PROVAMT,ED_TD1DOL,ED_EMPTYPE,ED_ELIGIBLE,ED_SECTION,ED_ORG,ED_DEPTNO,ED_DEPTEDATE,ED_DIVEDATE,ED_GLNO,ED_LOC,ED_ADMINBY,ED_REGION,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & EID&, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
'Ticket #23564 Franks 04/15/2013
rsEmp.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & EID&, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
If rsEmp.EOF Then Exit Function

If glbLinamar Then
    If lblCountry.Visible Then
        rsEmp("ED_COUNTRY") = Replace(Mid(lblCountry, InStr(lblCountry, "to """) + 3), """", "")
    End If
    If lblProv.Visible Then
        rsEmp("ED_PROV") = Replace(Mid(lblProv, InStr(lblProv, "to """) + 3), """", "")
    End If
    rsEmp("ED_SIN") = medSIN
    rsEmp("ED_DIV") = Right(EID&, 3)
End If

rsEmp("ED_DOH") = dlpDOH

'Ticket #29007 - City of Campbell River do not want auto update the Seniority Date. Also not transferring the Seniority Date to Vadim on Rehire
'Ticket #21935 - Do not change the ED_SENDTE for Friesens as it is Original Hire date for them (relabeled)
If glbCompSerial <> "S/N - 2279W" And glbCompSerial <> "S/N - 2458W" Then
    rsEmp("ED_SENDTE") = dlpDOH
End If

'Hemu - 07/21/04 Begin - County of Essex - Modifications  - Ticket # 6549
'Musashi - Ticket #15310
If glbCompSerial = "S/N - 2192W" Or glbCompSerial = "S/N - 2380W" Or glbCompSerial = "S/N - 2288W" Then
    'Ticket #29984 - County of Essex - They want to change from Last Hire to Union Date
    If glbCompSerial = "S/N - 2192W" Then
        rsEmp("ED_UNION") = dlpLTHire
    Else
        rsEmp("ED_LTHIRE") = dlpLTHire
    End If
End If

If glbVadim Then
    rsEmp("ED_FDAY") = dlpLTHire
    
    'Ticket #19113 - District Municipality of Muskoka
    If glbCompSerial = "S/N - 2373W" Then
        'Payroll ID same as Employee #
        If Len(rsEmp("ED_PAYROLL_ID")) < 1 Or IsNull(rsEmp("ED_PAYROLL_ID")) Then
            'Payroll ID same as Employee #
            rsEmp("ED_PAYROLL_ID") = EID&
        End If
    End If
    
ElseIf IsDate(dlpLTHire) Then   'Ticket #18668
    'Ticket #29984 - County of Essex - Not for them
    If glbCompSerial = "S/N - 2192W" Then
        'Do nothing for now
    Else
        rsEmp("ED_LTHIRE") = dlpLTHire
    End If
End If
'Hemu 07/21/04 End

'City of Timmins - Ticket #11066 - They want to reset the TD1 Dollar and Prov Amount
                    'based on the Company Master for Rehires
If glbCompSerial = "S/N - 2375W" Then
    rsPA.Open "select PC_NEXT_AVAILABLE_NBR,PC_FEDTAX,PC_PROVTAX from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsPA.EOF Then
        rsEmp("ED_TD1DOL") = rsPA("PC_FEDTAX")
        rsEmp("ED_PROVAMT") = rsPA("PC_PROVTAX")
    End If
    rsPA.Close
End If

If glbWFC Then
    If Left(comEmpType.Text, 1) = "Y" Then
        rsEmp("ED_EMPTYPE") = "Y"
        'Rehire: When processing a rehire, the Membership Entry date is being updated with the first of the month following.
        'This logic only exists for Pension Type DBKITCH. All other Pension Types should have their entry date equal to the DOH
        If getDBType(rsEmp("ED_SECTION"), rsEmp("ED_ORG"), "PenType") = "DBKITCH" Then
            If IsDate(dlpDOH.Text) Then
                xTDate = CVDate(dlpDOH.Text)
                xTDate = DateAdd("M", 1, xTDate)
                xTDate = CVDate(MonthName(month(xTDate)) & " 1," & Str(Year(xTDate)))
                rsEmp("ED_ELIGIBLE") = xTDate
                rsEmp.Update
            End If
        Else
            If IsDate(dlpDOH.Text) Then
                xTDate = CVDate(dlpDOH.Text)
                rsEmp("ED_ELIGIBLE") = xTDate
                rsEmp.Update
            End If
        End If
        Call Pause(1)
        Call WFCPensionMaster(EID&, , "H", , Year(dlpDOH.Text), "Rehire")
    End If
    If Left(comEmpType.Text, 1) = "N" Then
        rsEmp("ED_EMPTYPE") = "N"
    End If
    'If rsEMP("ED_WORKCOUNTRY") = "U.S.A." Then 'Ticket #23564 Franks 04/15/2013
        If IsDate(dlpDOH.Text) Then rsEmp("ED_PTEDATE") = CVDate(dlpDOH.Text)
    'End If
    If IsWFCNGSEmployee = True Then  'Ticket #25521 Franks 06/03/2014
        rsEmp("ED_SMOKER") = 1 '"   Smoker should be Y.
        rsEmp("ED_TYPEVEHICLE") = Null '"   Remove any smoker data in type of vehicle and permit #2.
        rsEmp("ED_PARKPERMIT2") = Null
    End If
End If

'Ticket #20570 - Jerry said to retain the original values if no new values entered
'on these fields except for Samuel.
'Ticket #19310 - Samuel, Son & Co., Limited
If glbCompSerial = "S/N - 2382W" Then
    'Update the rehired record with new values entered on the rehire screen
    rsEmp("ED_DEPTNO") = clpDept.Text
    If IsDate(dlpDeptEDate) Then
        rsEmp("ED_DEPTEDATE") = dlpDeptEDate
    End If
    If Len(clpGLNum.Text) > 0 Then
        rsEmp("ED_GLNO") = clpGLNum.Text
    Else
        rsEmp("ED_GLNO") = Null
    End If
    rsEmp("ED_DIV") = clpDiv1.Text
    If IsDate(dlpDivEDate) Then
        rsEmp("ED_DIVEDATE") = dlpDivEDate
    End If
    rsEmp("ED_LOC") = clpCode(1).Text
    rsEmp("ED_ADMINBY") = clpCode(3).Text
    rsEmp("ED_REGION") = clpCode(2).Text
    rsEmp("ED_SECTION") = clpCode(4).Text
    rsEmp("ED_EMP") = "A" 'Ticket #20648 Franks 09/26/2011
Else
    'Update the rehired record with new values entered on the rehire screen
    If Len(clpDept.Text) > 0 Then
        rsEmp("ED_DEPTNO") = clpDept.Text
    End If
    If IsDate(dlpDeptEDate) Then
        rsEmp("ED_DEPTEDATE") = dlpDeptEDate
    End If
    If Len(clpGLNum.Text) > 0 Then
        rsEmp("ED_GLNO") = clpGLNum.Text
    End If
    If Len(clpDiv1.Text) > 0 Then
        rsEmp("ED_DIV") = clpDiv1.Text
    End If
    If IsDate(dlpDivEDate) Then
        rsEmp("ED_DIVEDATE") = dlpDivEDate
    End If
    If Len(clpCode(1).Text) > 0 Then
        rsEmp("ED_LOC") = clpCode(1).Text
    End If
    If Len(clpCode(3).Text) > 0 Then
        rsEmp("ED_ADMINBY") = clpCode(3).Text
    End If
    If Len(clpCode(2).Text) > 0 Then
        rsEmp("ED_REGION") = clpCode(2).Text
    End If
    If Len(clpCode(4).Text) > 0 Then
        rsEmp("ED_SECTION") = clpCode(4).Text
    End If
    If glbWFC Then 'Ticket #24582 Franks 11/13/2013
        'o   Date range from the termination record should not be restored. The From Date should equal the Rehire Date. No To Date is required.
        If IsDate((dlpDOH.Text)) Then rsEmp("ED_SFDATE") = CVDate(dlpDOH.Text) Else rsEmp("ED_SFDATE") = Null
        rsEmp("ED_STDATE") = Null
    End If
End If

If Len(clpCode(0).Text) > 0 Then rsEmp("ED_EMP") = clpCode(0).Text 'Ticket #23820 Franks 05/29/2013
If Len(clpCode(5).Text) > 0 Then rsEmp("ED_ORG") = clpCode(5).Text 'Ticket #23820 Franks 05/29/2013
If Len(clpPT.Text) > 0 Then rsEmp("ED_PT") = clpPT.Text 'Ticket #25562 Franks 06/17/2014

If glbWFC Then 'Ticket #23247 Franks 09/17/2013 - US NGS/Ben
    Call DispNGSBenGroups
    If Len(NewBGroup) > 0 Then rsEmp("ED_BENEFIT_GROUP") = NewBGroup Else rsEmp("ED_BENEFIT_GROUP") = Null
    If Len(NewPayGroup) > 0 Then rsEmp("ED_VADIM2") = NewPayGroup
    glbWFCPayGroup = NewPayGroup
    If Len(NewNGSSub) > 0 Then rsEmp("ED_VADIM1") = NewNGSSub Else rsEmp("ED_VADIM1") = Null
    glbWFCNGSSubGroup = NewNGSSub
    If glbCandidate > 0 And xHRSoftUpt Then 'Ticket #24184 Franks 09/25/2013
        rsEmp("ED_CANDIDATE") = glbCandidate
        Call WFCHRSoftProcUpt("frmEREHIRE")
    End If
End If

'Ticket #27841 - Decor Cabinet - Change the Entitlement Period based on the new Hire Date.
'They do Anniversary Month Year End
If glbCompSerial = "S/N - 2451W" Then
    If IsDate(dlpDOH.Text) Then
        If Year(dlpDOH.Text) = Year(Now) Then
            rsEmp("ED_EFDATE") = dlpDOH.Text
            rsEmp("ED_ETDATE") = DateAdd("yyyy", 1, CVDate(dlpDOH.Text) - 1)
        Else
            rsEmp("ED_EFDATE") = Null
        End If
    End If
End If

'WDGPHU - Ticket #27899
If glbCompSerial = "S/N - 2411W" Then
    'Multi-Position Employee
    rsEmp("ED_ORGT1") = "NO"
End If

'Ticket #25412 - Town of Greater Napanee
If glbCompSerial = "S/N - 2447W" Then
    'EI Start Date - Relabelled field update with Original Hire Date
    rsEmp("ED_UNION") = dlpDOH
End If

rsEmp.Update

'Ticket #30412 - If the Date of Rehire(Hire) is not this year then let the system compute the new entitlement period based on the new rehire(hire).
If glbCompSerial = "S/N - 2451W" Then
    Dim SQLQ As String
    If Year(dlpDOH.Text) <> Year(Now) Then
        SQLQ = "ED_EMPNBR = " & rsEmp("ED_EMPNBR")
        Call EntReCalc(SQLQ)
    End If
End If

Reset_BASIC = True

Screen.MousePointer = DEFAULT

End Function

Private Sub lblEENumNew_Change()
txtEmpNo = lblEENumNew
End Sub

Private Sub txtEmpID_Change()
Call CountEmpNbr
End Sub
Private Sub clpDiv_LostFocus()
If Len(clpDiv) > 0 And clpDiv.Caption <> "Unassigned" Then
    If glbLinamar Then Call SetCountry(clpDiv)
    Call CountEmpNbr
End If

End Sub

Private Sub SetCountry(wDIV)
Dim oCountry, oProv
Dim wCountry, wProv
' danielk - 12/31/2002 - changed all references to rsDATA2 to rsDATA2
oCountry = rsDATA2("ED_COUNTRY")
oProv = rsDATA2("ED_PROV")
Select Case wDIV
Case "501", "505"
    wCountry = "MEXICO"
    wProv = "CO"
Case "502"
    wCountry = "MEXICO"
    wProv = "DU"
Case "117"
    wCountry = "U.S.A."
    wProv = "IA"
Case "118"
    wCountry = "U.S.A."
    wProv = "IL"
Case "120"
    wCountry = "U.S.A."
    wProv = "IA"
Case "345", "310", "370"
    wCountry = "U.S.A."
    wProv = oProv
Case "360"
    wCountry = "U.S.A."
    wProv = "KY"
Case "430", "440"
    wCountry = "GERMANY"
    wProv = oProv
Case Else
    wCountry = "CANADA"
    wProv = "ON"
End Select
'''Ticket #22819 Franks 11/23/2012 - begin
''If oCountry <> wCountry Then
''    lblCountry = "Country will change from """ & oCountry & """ to """ & wCountry & """"
''    lblCountry.Visible = True
''
''    fraBasicInfo.Top = 7080
''Else
''    lblCountry = ""
''    lblCountry.Visible = False
''
''    fraBasicInfo.Top = 6580
''End If
''If oProv <> wProv Then
''    lblPROV = "Province will change from """ & oProv & """ to """ & wProv & """"
''    lblPROV.Visible = True
''
''    fraBasicInfo.Top = 7080
''Else
''    lblPROV = ""
''    lblPROV.Visible = False
''
''    fraBasicInfo.Top = 6580
''End If
''If gSec_Show_SIN_SSN Then
''    If wCountry = "CANADA" Then
''        If Not IsNull(rsDATA2("ED_SIN")) Then
''            If Not SIN_chk(rsDATA2("ED_SIN")) Then
''                lblSIN.Caption = "Enter S.I.N."
''                lblSIN.Visible = True
''                medSIN.Tag = "11-Social Insurance Number"   '
''                medSIN.Mask = "###-###-###"
''                medSIN.Visible = True
''            End If
''        End If
''    End If
''End If

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
UpdateRight = gSec_Upd_Rehire
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = xUpdateable
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
'Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim EID&
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    If Not UserEmailExist Then
        Exit Sub
    End If

    'Ticket #18090 - begin
    If Len(txtEmpNo) = 0 Then
        EID& = lblEENum
    Else
        EID& = getEmpnbr(txtEmpNo)
    End If
    xToEmail = GetComPreferEmail("EMAIL_ONREHIRE", EID&)
    If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
        xToEmail = GetComPreferEmail("EMAIL_ONREHIRE")
    End If
    'Ticket #18090 - end
    If Len(xToEmail) > 0 Then
        frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONREHIRE")
        'frmSendEmail.txtSubject.Text = "info:HR Employee Rehire Notice"
        'Ticket #18578
        'frmSendEmail.txtSubject.Text = "info:HR Employee Rehire Notice - " & lblEEName.Caption
        'Ticket #18755
        xBranch = GetEmpData(EID&, "ED_SECTION", "")
        If Len(xBranch) > 0 Then
            xBranch = xBranch & " - "
        End If
        xEmailSubject = "info:HR Employee Rehire Notice - " & xBranch & lblEEName.Caption
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
Dim EID&
On Error GoTo Email_Err
        If Not UserEmailExist Then
            Exit Sub
        End If

        'Ticket #20317 - More Emails for everyone
        xToEmail = GetComPreferEmail("EMAIL_ONREHIRE", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONREHIRE")
        End If
            
        frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONREHIRE")
        'frmSendEmail.txtCC.Text = xEmail
        'frmSendEmail.txtSubject.Text = "info:HR Employee Rehire Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Employee Rehire Notice - " & lblEEName.Caption
        frmSendEmail.txtBody.Text = MailBody
        frmSendEmail.Show 1

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

Private Sub UpdBenEffectiveDate(xEESEQ, xDOH)
Dim rsTBen As New ADODB.Recordset
Dim SQLQ As String, xEDate
    If Not IsDate(xDOH) Then Exit Sub
    SQLQ = "SELECT * FROM Term_HRBENFT WHERE (Term_HRBENFT.TERM_SEQ= " & xEESEQ & " )"
    rsTBen.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
    Do While Not rsTBen.EOF
        xEDate = xDOH
        If Not IsNull(rsTBen("BF_WaitPeriod")) Then
            If IsNumeric(rsTBen("BF_WaitPeriod")) Then
                If Not IsNull(rsTBen("BF_DWM")) Then
                    If UCase(rsTBen("BF_DWM")) = "D" Then
                        xEDate = DateAdd("D", rsTBen("BF_WaitPeriod"), xEDate)
                    End If
                    If UCase(rsTBen("BF_DWM")) = "W" Then
                        xEDate = DateAdd("WW", rsTBen("BF_WaitPeriod"), xEDate)
                    End If
                    If UCase(rsTBen("BF_DWM")) = "M" Then
                        xEDate = DateAdd("M", rsTBen("BF_WaitPeriod"), xEDate)
                    End If
                End If
            End If
        End If
        rsTBen("BF_EDATE") = xEDate
        rsTBen.Update
        rsTBen.MoveNext
    Loop
    rsTBen.Close
End Sub

Private Sub ComEType()
comEmpType.Clear
comEmpType.AddItem "0 - Not Applicable"
comEmpType.AddItem "1 - Full Time Salary"
'If glbCompSerial <> "S/N - 2380W" Then   'VitalAire Canada Ticket #14736
    comEmpType.AddItem "2 - Part Time Salary"
'End If
comEmpType.AddItem "3 - Full Time Hourly"
comEmpType.AddItem "4 - Part Time Hourly"
comEmpType.AddItem "5 - Casual/Other"
comEmpType.AddItem "6 - Contract Salary"
comEmpType.AddItem "7 - Contract Hourly"    '23June99 js
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14827
    'comEmpType.AddItem "8 - 80% Full Time Salary"
    comEmpType.AddItem "8 - 80% @7.5hrs"   'Ticket #14995
Else
    comEmpType.AddItem "8 - Salary Pensioners"
End If
'comEmpType.AddItem "9 - Salary Elected officials"
'Added by Bryan 12/08/05 for granite club
If glbCompSerial = "S/N - 2241W" Then
    comEmpType.AddItem "9 - Commissioned Employees"
ElseIf glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14827
    comEmpType.AddItem "9 - Former Air Liquide Emp."
Else
    comEmpType.AddItem "9 - Salary Elected officials"
End If
If glbCompSerial = "S/N - 2380W" Then   'VitalAire Canada Ticket #14736
    'Ticket #14827
    comEmpType.AddItem "A - 10hr day @86.67"
    comEmpType.AddItem "B - 12 hour day @78"
    comEmpType.AddItem "C - 80% @6hrs"      ''Ticket #14995
    'comEmpType.AddItem "10 - 10hr day @86.67"
    'comEmpType.AddItem "11 - 12 hour day @78"
End If

'Ticket# 10189
If glbCompSerial = "S/N - 2214W" Then 'Casey House Hospice
    comEmpType.Clear
    comEmpType.AddItem "0 - Not Applicable"
    comEmpType.AddItem "1 - Full Time"
    comEmpType.AddItem "2 - Part Time Regular"
    comEmpType.AddItem "3 - Part Time Temporary Full Time"
    comEmpType.AddItem "4 - Part Time Job Share"
    comEmpType.AddItem "5 - Casual Regular"
    comEmpType.AddItem "6 - Casual Temporary Full Time"
End If
'Ticket #14752
If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab Hospital
    comEmpType.Clear
    comEmpType.AddItem "A - Temp FT"
    comEmpType.AddItem "B - Temp PT"
    comEmpType.AddItem "C - Casual"
    comEmpType.AddItem "F - FT"
    comEmpType.AddItem "J - Job Share"
    comEmpType.AddItem "P - PT"
    comEmpType.AddItem "S - Student"
    comEmpType.AddItem "X - Terminated"
End If
'Ticket #15794
If glbCompSerial = "S/N - 2390W" Then 'Collectcorp
    comEmpType.Clear
    comEmpType.AddItem "0 - No"
    comEmpType.AddItem "1 - Yes"
End If

'Ticket #16889
If glbCompSerial = "S/N - 2384W" Then 'Town of St. Marys
    comEmpType.Clear
    comEmpType.AddItem "1 - Hourly"
    comEmpType.AddItem "2 - Salary"
    comEmpType.AddItem "3 - Volunteer"
    comEmpType.AddItem "4 - Elected Official"
End If

'Ticket #17077
If glbCompSerial = "S/N - 2172W" Then 'County of Lanark
    comEmpType.Clear
    comEmpType.AddItem "C - Part Time Temp"
    comEmpType.AddItem "F - Full Time Regular"
    comEmpType.AddItem "P - Part Time Regular"
    comEmpType.AddItem "T - Full Time Temp"
    comEmpType.AddItem "O - Other" 'Ticket #17076
End If

'Ticket #16395
If glbWFC Then
    comEmpType.Clear
    comEmpType.AddItem "Y - Yes"
    comEmpType.AddItem "N - No"
End If

End Sub

Private Sub UptTaxExampt()
Dim rsPARCO As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRPARCO"
    rsPARCO.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPARCO.EOF Then
        If Not IsNull(rsPARCO("PC_FEDTAX")) Then
            rsDATA2("ED_TD1DOL") = rsPARCO("PC_FEDTAX")
        End If
        If Not IsNull(rsPARCO("PC_PROVTAX")) Then
            rsDATA2("ED_PROVAMT") = rsPARCO("PC_PROVTAX")
        End If
    End If
    rsPARCO.Close
End Sub
Private Sub SamuelDefaultDat(xEmpNo, xTermSEQ)
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate

    SQLQ = "SELECT * FROM Term_HREMP WHERE TERM_SEQ = " & xTermSEQ & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If Not IsNull(rsEmpee("ED_DEPTNO")) Then clpDept.Text = rsEmpee("ED_DEPTNO")
        If Not IsNull(rsEmpee("ED_GLNO")) Then clpGLNum.Text = rsEmpee("ED_GLNO")
        If Not IsNull(rsEmpee("ED_DIV")) Then clpDiv1.Text = rsEmpee("ED_DIV")
        If Not IsNull(rsEmpee("ED_LOC")) Then clpCode(1).Text = rsEmpee("ED_LOC")
        If Not IsNull(rsEmpee("ED_ADMINBY")) Then clpCode(3).Text = rsEmpee("ED_ADMINBY")
        If Not IsNull(rsEmpee("ED_REGION")) Then clpCode(2).Text = rsEmpee("ED_REGION")
        If Not IsNull(rsEmpee("ED_SECTION")) Then clpCode(4).Text = rsEmpee("ED_SECTION")
        If Not IsNull(rsEmpee("ED_DEPTEDATE")) Then dlpDeptEDate.Text = rsEmpee("ED_DEPTEDATE")
        If Not IsNull(rsEmpee("ED_DIVEDATE")) Then dlpDivEDate.Text = rsEmpee("ED_DIVEDATE")
        If Not IsNull(rsEmpee("ED_LTHIRE")) Then dlpLTHire.Text = rsEmpee("ED_LTHIRE")
        'Ticket #23820 Franks 05/29/2013 - begin
        If lblEEStatus.Visible Then
            If Not IsNull(rsEmpee("ED_EMP")) Then clpCode(0).Text = rsEmpee("ED_EMP")
        End If
        If lblUnion.Visible Then
            If Not IsNull(rsEmpee("ED_ORG")) Then clpCode(5).Text = rsEmpee("ED_ORG")
        End If
        'Ticket #23820 Franks 05/29/2013 - end
    End If
    rsEmpee.Close
    
End Sub
Private Sub WFCOther2Screen(xEmpNo, xTermSEQ)
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
    
    SaveGLNo = ""
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    lbOtherDate1.Visible = False
    dlpDOther1.Visible = False
    
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2, ED_GLNO FROM Term_HREMP WHERE TERM_SEQ = " & xTermSEQ & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
        If IsNull(rsEmpee("ED_GLNO")) Then SaveGLNo = "" Else SaveGLNo = rsEmpee("ED_GLNO")
        clpDiv1.Text = glbEmpDiv 'Ticket #24582 Franks 11/11/2013
    End If
    rsEmpee.Close
    
    Call WFC_Disp_NGSStartDate(glbEmpDiv) 'Ticket #24652 Franks 12/02/2013
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    
    xNGSStart = ""
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM Term_HREMP_OTHER WHERE TERM_SEQ = " & xTermSEQ & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close

    'Ticket #20385 Franks 05/31/2011
    ''No NGS Effective Date, skip
    'If Len(xNGSStart) = 0 Then Exit Sub
    
    '''Ticket #24652 Franks 12/02/2013 - move to WFC_Disp_NGSStartDate
    ''lbOtherDate1.Caption = lStr("Other Date 1")
    ''lbOtherDate1.Top = lblEEType.Top + 330
    ''dlpDOther1.Top = lblEEType.Top + 330
    ''lbOtherDate1.Visible = True
    ''dlpDOther1.Visible = True

End Sub

Private Sub WFC_Disp_NGSStartDate(xDiv) 'Ticket #24652 Franks 12/02/2013
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
    IsWFCNGSEmployee = False
    If Len(xDiv) > 0 Then
        '"   If the Division is a NGS Division, the NGS Start Date should be mandatory.
        SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDiv & "' "
        'SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
        'SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & xStatus & "' "
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            lbOtherDate1.Caption = lStr("Other Date 1")
            lbOtherDate1.Top = lblEEType.Top + 330
            dlpDOther1.Top = lblEEType.Top + 330
            lbOtherDate1.FontBold = True
            If Len(clpPT.Text) > 0 Then 'Ticket #25562 Franks 06/17/2014
                If Not clpPT.Text = "FT" Then
                    lbOtherDate1.FontBold = False
                End If
            End If
            lbOtherDate1.Visible = True
            dlpDOther1.Visible = True
            IsWFCNGSEmployee = True 'Ticket #25521 Franks 06/03/2014
        Else
            lbOtherDate1.FontBold = False
            lbOtherDate1.Visible = False
            dlpDOther1.Visible = False
        End If
    End If
End Sub

Private Sub WFC_NGS_Trans(xEmpNo)
Dim xLDate
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE1", Null)
    Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE2", Null)
    Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE3", Null) 'Ticket #25604 Franks 06/18/2014
    If IsDate(dlpDOther1.Text) Then
        xLDate = dlpDOther1.Text 'Date
        Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE1", CVDate(dlpDOther1.Text))
        'Ticket #25604 Franks 06/18/2014 - On Rehire - US Employee, Other Date 6 should equal Other Date 1
        Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE6", CVDate(dlpDOther1.Text))
        Call NGSAuditAdd(xEmpNo, "M", "Rehire", lStr("Other Date 1"), "", (dlpDOther1.Text), xLDate)
    End If
End Sub

Private Sub EEO_Process(xEmpNo) 'Ticket #20270 Franks 05/05/2011
    If glbEmpCountry = "U.S.A." Then
        'Call uptEEO_Fields(xEmpNo, "New")
        'Ticket #25669 Franks 06/24/2014
        Call uptEEO_Fields(xEmpNo, "Update")
    End If
End Sub

Private Sub AUDIT_GWL_TRANS(xEmpNo)
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpID
Dim xForm As String
Dim xTranType
Dim xChgType
Dim xEDate, xDate1, xDate2
Dim xLDate
Dim xBenGroup

On Error GoTo AUDIT_ERR

    If Not glbIsGWL Then Exit Sub
    SQLQ = "SELECT ED_EMPNBR, ED_BENEFIT_GROUP, ED_DOH, ED_DIV, ED_ORG FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_BENEFIT_GROUP")) Then xBenGroup = "" Else xBenGroup = rsEmpee("ED_BENEFIT_GROUP")
    End If
    'rsEmpee.Close
    'No Benefit Group Code, skip
    If Len(xBenGroup) = 0 Then Exit Sub
    xEmpID = xEmpNo
    xTranType = "U"
    xChgType = "Rehire"

    'If IsDate(dlpBenCeaseDate.Text) Then
    '    xEDate = dlpBenCeaseDate.Text
    'Else
    '    xEDate = dlpTermDate.Text
    'End If
    xEDate = Date
    xForm = "Re-hire"
    xLDate = Date
    If CVDate(xEDate) > CVDate(xLDate) Then
        xLDate = xEDate
    End If
    rsEmpee.Close
    
    'GWL field changes --------------------------------------
    Call GWLAuditAdd(xEmpID, xTranType, xChgType, xEDate, xForm, "New Hire Date", "", dlpDOH.Text, xLDate)

    Exit Sub

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING GWL AUDIT RECORD", "GWL AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub

Private Sub setEffDates(xDOH) 'Ticket #23837 Franks 05/28/2013
    If Not glbSamuel Then
        dlpDeptEDate.Text = xDOH
        dlpDivEDate.Text = xDOH
    End If
End Sub

Private Sub showStatusUnionFields() 'Ticket #23820 Franks 05/28/2013
    lblEEStatus.Visible = True
    clpCode(0).Visible = True
    lblUnion.Caption = "New " & lStr("Union")
    lblUnion.Visible = True
    clpCode(5).Visible = True
    'Ticket #25562 Franks 06/17/2014 - begin
    lblPT.Caption = "New " & lStr("Category")
    lblPT.Visible = True
    clpPT.Visible = True
    'Ticket #25562 Franks 06/17/2014 - end
    If glbWFC Then
        lblEEStatus.FontBold = True
        lblUnion.FontBold = True
        lblTitle(13).FontBold = True
        lblPT.FontBold = True
    End If
End Sub

Private Sub WFCMainScreen()
    lblTitle(7).Visible = True
    txtPayrollID.Visible = True
    lblTitle(7).Top = 4780
    txtPayrollID.Top = 4780
    lblPayIDExist.Top = 4780
    'Ticket #16395 - begin
    lblEEType.Top = 5130
    comEmpType.Top = 5130
    lblEEType.Visible = True '
    comEmpType.Visible = True
    'Ticket #16395 - end
    
    'Ticket #19266 Franks 12/02/10
    'On Rehire screen, add "Other Date 1" to be completed by the user. This is an optional field.
    'If glbWFC Then
        Call WFCOther2Screen(glbTERM_ID, glbTERM_Seq)
        Call showStatusUnionFields 'Ticket #23820 Franks 05/28/2013
    'End If
End Sub

Private Sub DispNGSBenGroups()  'Ticket #23247 Franks 09/13/2013
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xLocOrg
Dim xDivCountry
Dim xUnion, xDiv, xStatus
    NewBGroup = ""
    NewPayGroup = " "
    NewNGSSub = ""
    If Len(clpDiv1.Text) = 0 Then Exit Sub
    xDiv = clpDiv1.Text
    If Len(clpCode(5).Text) = 0 Then Exit Sub 'Union
    xUnion = clpCode(5).Text
    If Len(clpCode(0).Text) = 0 Then Exit Sub 'Status
    xStatus = clpCode(0).Text
    
    xDivCountry = GetCountryFromDiv(clpDiv1.Text)
    If Not xDivCountry = "U.S.A." Then
        Exit Sub
    End If

    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDiv & "' "
    SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
    SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & xStatus & "' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If rsTemp.EOF Then 'Ticket #23564 Franks 04/17/2013
    'check "-" status, such as "-ACT2", convert "-ACT2" to "ACT2" then compare ED_EMP with not equal to
        SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDiv & "' "
        SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
        SQLQ = SQLQ & "AND LEFT(NG_PLAN_CODE,1) = '-' " 'for "-" code only
        SQLQ = SQLQ & "AND NOT ((CASE LEFT(NG_PLAN_CODE,1) WHEN '-' THEN REPLACE(NG_PLAN_CODE,'-', '') ELSE '' END) = '" & xStatus & "') " 'convert "-ACT2" to "ACT2"; no "-" then ""
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        'if not found then without Status code
        If rsTemp.EOF Then
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDiv & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
            SQLQ = SQLQ & "AND ((NG_PLAN_CODE IS NULL) OR NOT( NG_PLAN_CODE ='" & xStatus & "')) "
            If rsTemp.State <> 0 Then rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        End If
    End If
    
    If Not rsTemp.EOF Then
        NewPayGroup = rsTemp("NG_PAY_GROUP")
        NewNGSSub = rsTemp("NG_SUB_GROUP")
        If Not IsNull(rsTemp("NG_BENEFIT_GROUP")) Then 'Ticket #23903 Franks 06/20/2013
            NewBGroup = rsTemp("NG_BENEFIT_GROUP")
        End If
    End If
    rsTemp.Close
        
    If Not clpPT.Text = "FT" Then 'Ticket #25562 Franks 06/17/2014
        '"   If Category not equal FT and it's a US employee, the NGS Start Date is not mandatory and there is no benefit group code
         NewBGroup = ""
    End If
    
End Sub

Private Sub locBeneGroupUpdate(xEmpNo) 'Ticket #23247 Franks 09/17/2013
'Dim NewBGroup
Dim Msg
Dim SQLQ
Dim rsEmp As New ADODB.Recordset
Dim rsBenT As New ADODB.Recordset
Dim xTemp
    
    If Not IsWFCUSBenEmp(xEmpNo) Then
        Exit Sub 'for US employees only for now
    End If
    
    'This function works for Benefit Group change, so the previous and current
    'benefit group can't be blank
    'NewBGroup = clpBGroup
    'If Len(SaveBGroup) = 0 Then Exit Sub
    If Len(NewBGroup) = 0 Then
        'Ticket #21677 Franks 09/16/2013 - begin
        '"1.If the Transfer In Division or Union is not a NGS location (ie: not found in the matrix):
        '"2.Delete the Benefit Group Code
        '"3. Remove the Benefit Group Code from the Company Paid Benefits
        '"   Add a Benefit End Date to equal the Transfer In Date minus 1.
        If Len(SaveBGroup) > 0 Then
            '1.
            SQLQ = "UPDATE HREMP SET ED_BENEFIT_GROUP = NULL WHERE ED_EMPNBR = " & xEmpNo & " "
            gdbAdoIhr001.Execute SQLQ
            '3.
            If IsDate(dlpDOH.Text) Then
                xTemp = DateAdd("D", -1, CVDate(dlpDOH.Text))
                'SQLQ = "UPDATE HRBENFT SET BF_CEASEDATE = " & Date_SQL(xTemp) & " WHERE BF_EMPNBR = " & xEmpNo & " "
                'SQLQ = SQLQ & "AND BF_GROUP = '" & SaveBGroup & "' "
                'SQLQ = SQLQ & "AND BF_PCC = 1 "
                'gdbAdoIhr001.Execute SQLQ
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND BF_GROUP = '" & SaveBGroup & "' "
                SQLQ = SQLQ & "AND BF_PCC = 1 "
                'If Not rsBenT.State <> 0 Then rsBenT.Close
                rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsBenT.EOF
                    rsBenT("BF_CEASEDATE") = CVDate(xTemp)
                    rsBenT.Update
                    'update audit
                    Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
                    rsBenT.MoveNext
                Loop
            End If
            '2.
                SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE BF_EMPNBR = " & xEmpNo & " "
                'SQLQ = SQLQ & "AND BF_GROUP = '" & SaveBGroup & "' "
                'SQLQ = SQLQ & "AND BF_PCC = 1 "
                gdbAdoIhr001.Execute SQLQ
        End If
        'Ticket #21677 Franks 09/16/2013 - end
        Exit Sub
    End If
    
    'If SaveBGroup <> NewBGroup Then
    If Len(NewBGroup) > 0 Then
        If IsWFCUSBenEmp(xEmpNo) Then 'Ticket #23247 Franks 09/16/2013
            If SaveBGroup <> NewBGroup Then
                Call WFC_UptUSBenByEmp(xEmpNo, CVDate(dlpDOH.Text), 0, "Y", "Y", , SaveBGroup, dlpDOH.Text)
            Else
                Call WFC_UptUSBenByEmp(xEmpNo, CVDate(dlpDOH.Text), 0, "Y", "Y", , , dlpDOH.Text)
            End If
            Exit Sub
        Else
            Exit Sub 'do not show pop screen for now
            Msg = "Do you want add/update the Employee's Benefits "
            Msg = Msg & " with the Benefit Codes defined for the Benefit Group? "
            If MsgBox(Msg, 36, "info:HR") = 6 Then
                'Call UpdateBenefitGroup
                Call glbUpdateBenefitGroup(xEmpNo, SaveBGroup, NewBGroup, dlpDOH.Text)
                DoEvents
                glbLEE_ID = xEmpNo
                'glbLEE_SName = locSurName
                'glbLEE_FName = locFName
                frmBENGRLIST.Show 1
            'Else 'Frank 10/04/2003 Delete Benefit Group on Employee Benefit screen if wipe off the Benefit Group
            '    If Len(clpBGroup.Text) = 0 Then
            '        SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE NOT (BF_GROUP IS NULL) AND BF_EMPNBR =" & xlocEmpnbr
            '        gdbAdoIhr001.Execute SQLQ
            '    End If
            End If
            
            'If the Benefit Group changes, go to the Benefit Group Matrix to update
            'the HREMP 's Coverage Class and Benefit Account on the status/dates screen.
            If Len(NewBGroup) > 0 Then
                ''SQLQ = "SELECT * FROM HREMP"
                ''SQLQ = SQLQ & " WHERE HREMP.ED_EMPNBR = '" & xEmpNo & "' "
                ''rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                ''If Not rsEmp.EOF Then
                ''    Call getValsFromBenGrpMatrix(NewBGroup, rsEmp("ED_DIV"))
                ''    If Len(xBenAccount) = 0 Then
                ''        rsEmp("ED_USER_NUM1") = Null
                ''    Else
                ''        rsEmp("ED_USER_NUM1") = xBenAccount
                ''    End If
                ''    If Len(xCovClass) = 0 Then
                ''        rsEmp("ED_USER_TEXT2") = Null
                ''    Else
                ''        rsEmp("ED_USER_TEXT2") = xCovClass
                ''    End If
                ''    rsEmp.Update
                ''End If
                ''rsEmp.Close
            End If
        End If
    End If

End Sub

Private Sub WFCHRSoftDispValues()
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xSF_STARTDATE
Dim xTemp

xHRSoftUpt = False
xSF_STARTDATE = ""
'xHRSoftPTCode = ""
If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    SQLQ = SQLQ & "AND SF_UPT_REHIRE = 0 "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsCanid.EOF Then
        xHRSoftUpt = True
        If Len(glbTrsEE_ID) > 0 Then txtEmpNo.Text = glbTrsEE_ID
        If Not IsNull(rsCanid("SF_STARTDATE")) Then
            dlpDOH.Text = rsCanid("SF_STARTDATE")
            dlpDeptEDate.Text = rsCanid("SF_STARTDATE")
            dlpDivEDate.Text = rsCanid("SF_STARTDATE")
            xSF_STARTDATE = rsCanid("SF_STARTDATE")
        End If
        If Not IsNull(rsCanid("SF_PAYROLL_ID")) Then txtPayrollID.Text = rsCanid("SF_PAYROLL_ID")
        If Not IsNull(rsCanid("SF_GLNO")) Then clpGLNum.Text = rsCanid("SF_GLNO")
        If Not IsNull(rsCanid("SF_LOC")) Then clpCode(1).Text = rsCanid("SF_LOC")
        If Not IsNull(rsCanid("SF_ADMINBY")) Then clpCode(3).Text = rsCanid("SF_ADMINBY")
        If Not IsNull(rsCanid("SF_REGION")) Then clpCode(2).Text = rsCanid("SF_REGION")
        If Not IsNull(rsCanid("SF_SECTION")) Then clpCode(4).Text = rsCanid("SF_SECTION")
        'Ticket #24451 Franks 10/18/2013
        'o   Bring the GL # in from term_HREMP and display it on the rehire screen
        If Len(SaveGLNo) > 0 Then clpGLNum.Text = SaveGLNo
        
        lblTitle(0).Caption = "Enter/Update " & """Original Hire Date"""
        lblEmpno.Caption = "Enter/Update Employee Number"
        chkRestore(9).Enabled = False
        chkRestore(10).Enabled = False
        chkRestore(16).Enabled = False
    End If
End If

If Len(glbTrsDept) > 0 Then clpDept.Text = glbTrsDept
If Len(glbTrsDIV) > 0 Then clpDiv1.Text = glbTrsDIV
If Len(glbTrsStatus) > 0 Then clpCode(0).Text = glbTrsStatus
If Len(glbTrsUnion) > 0 Then clpCode(5).Text = glbTrsUnion
    
'Ticket #24652 Franks 12/02/2013
'Call WFC_Disp_NGSStartDate(clpDIV1.Text)
If lbOtherDate1.Visible Then
    dlpDOther1.Text = xSF_STARTDATE
End If
'Ticket #24652 Franks 12/02/2013 - end

End Sub

Private Sub locHRSoftAction()
'o   On save, process a mini-new hire. We want the Banking Information screen to appear next followed by the Position and Salary
'the mini-new hire includes Demo, Status/Date, Position, Salary, Banking
'call public sub from wfc module
Dim rsCanid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTemp

'''xHRSoftUpt = False
'xHRSoftPTCode = ""
If glbCandidate > 0 Then
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & glbCandidate & " "
    If rsCanid.State <> 0 Then rsCanid.Close
    rsCanid.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCanid.EOF Then
        rsCanid("SF_UPT_REHIRE") = 1
        rsCanid.Update
        Call HRSoftAction(rsCanid)
    End If
    rsCanid.Close
End If

End Sub

Private Sub FamilDayEmpUpt(xEmpNo) 'Family Day Ticket #24729 01/21/2014 Franks
'"   On Rehire, if the employee being rehired has a Badge ID, the badge ID and the User Text 2 will be deleted during the rehire process.
Dim SQLQ As String
    SQLQ = "UPDATE HREMP SET ED_BADGEID = NULL, ED_USER_TEXT2 = NULL WHERE ED_EMPNBR = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub UptMissTroyNetworkLogin(xEmpNo) 'Ticket #28772 Franks 06/22/2016
Dim rsEmp As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_SECTION")) Then
            If rsEmp("ED_SECTION") = "MISS" Or rsEmp("ED_SECTION") = "TROY" Then
                SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
                If rsEmpOther.State <> 0 Then rsEmpOther.Close
                rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If rsEmpOther.EOF Then
                    rsEmpOther.AddNew
                    rsEmpOther("ER_COMPNO") = "001"
                    rsEmpOther("ER_EMPNBR") = xEmpNo
                    rsEmpOther("ER_NETWORKLOGIN") = getWFCNetworkLogin(rsEmp("ED_FNAME"), rsEmp("ED_SURNAME"))
                    rsEmpOther("ER_VENDORNO") = "N/A"
                Else
                    If IsNull(rsEmpOther("ER_NETWORKLOGIN")) Then
                        rsEmpOther("ER_NETWORKLOGIN") = getWFCNetworkLogin(rsEmp("ED_FNAME"), rsEmp("ED_SURNAME"))
                    Else
                        If Len(rsEmpOther("ER_NETWORKLOGIN")) = 0 Then
                            rsEmpOther("ER_NETWORKLOGIN") = getWFCNetworkLogin(rsEmp("ED_FNAME"), rsEmp("ED_SURNAME"))
                        End If
                    End If
                End If
                rsEmpOther("ER_LDATE") = Date
                rsEmpOther("ER_LTIME") = Time$
                rsEmpOther("ER_LUSER") = glbUserID
                rsEmpOther.Update
            End If
        End If
    End If
End Sub
