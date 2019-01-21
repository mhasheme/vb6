VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSHrsEnt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Hourly Entitlement Master"
   ClientHeight    =   12930
   ClientLeft      =   2565
   ClientTop       =   525
   ClientWidth     =   12660
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
   ScaleHeight     =   12930
   ScaleWidth      =   12660
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   6405
      LargeChange     =   315
      Left            =   10920
      Max             =   100
      SmallChange     =   315
      TabIndex        =   220
      Top             =   4320
      Width           =   300
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   765
      Left            =   0
      TabIndex        =   11
      Top             =   12165
      Width           =   12660
      _Version        =   65536
      _ExtentX        =   22331
      _ExtentY        =   1349
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
      Begin VB.CommandButton cmdDeletePrv 
         Appearance      =   0  'Flat
         Caption         =   "Delete Previous Year"
         Height          =   375
         Left            =   10080
         TabIndex        =   219
         Tag             =   "Delete all matching records to the above"
         Top             =   120
         Width           =   2145
      End
      Begin VB.CommandButton cmdDeleteAll 
         Appearance      =   0  'Flat
         Caption         =   "Delete All"
         Height          =   375
         Left            =   8160
         TabIndex        =   218
         Tag             =   "Delete all matching records to the above"
         Top             =   120
         Width           =   1785
      End
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   6360
         TabIndex        =   217
         Top             =   120
         Width           =   1665
      End
      Begin VB.CommandButton cmdDeleteEnt 
         Appearance      =   0  'Flat
         Caption         =   "&Delete Entitlement"
         Height          =   375
         Left            =   2520
         TabIndex        =   216
         Tag             =   "Delete all matching records to the above"
         Top             =   120
         Width           =   1785
      End
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "Print &All"
         Height          =   375
         Left            =   4440
         TabIndex        =   214
         Tag             =   "Print all Vacation Entitlement Report"
         Top             =   120
         Width           =   1785
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
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update Entitlement"
         Height          =   375
         Left            =   600
         TabIndex        =   215
         Tag             =   "Change all matching records to the above"
         Top             =   120
         Width           =   1785
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
   Begin VB.Frame VacFram03 
      BorderStyle     =   0  'None
      Height          =   4875
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11415
      Begin INFOHR_Controls.CodeLookup clpProv 
         Height          =   285
         Left            =   2520
         TabIndex        =   228
         Tag             =   "31-Province - Code"
         Top             =   2835
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   4
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fsHrsEnt.frx":0000
         Height          =   1455
         Left            =   0
         OleObjectBlob   =   "fsHrsEnt.frx":0014
         TabIndex        =   0
         Top             =   120
         Width           =   11175
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Left            =   3150
         TabIndex        =   7
         Tag             =   "40-As of Date"
         Top             =   3147
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Left            =   1340
         TabIndex        =   6
         Tag             =   "40-As of Date"
         Top             =   3147
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   6720
         TabIndex        =   3
         Tag             =   "01-Entitlement Code"
         Top             =   2511
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin Threed.SSFrame frmType 
         Height          =   375
         Left            =   1650
         TabIndex        =   224
         Top             =   3780
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin VB.TextBox txtUpdMethod 
            Appearance      =   0  'Flat
            DataField       =   "EH_UPDMETHOD"
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
            Left            =   3240
            MaxLength       =   1
            TabIndex        =   239
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
         Begin Threed.SSOption Replace 
            Height          =   195
            Left            =   2040
            TabIndex        =   10
            Tag             =   "Replace Entitlement Amount"
            Top             =   120
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Replace"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption Accum 
            Height          =   195
            Left            =   210
            TabIndex        =   9
            TabStop         =   0   'False
            Tag             =   "Add to Exist Entitlements"
            Top             =   120
            Width           =   1335
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Accumulate"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1335
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   2511
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1335
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   2193
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   6720
         TabIndex        =   5
         Tag             =   "00-Section - Code"
         Top             =   2829
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1335
         TabIndex        =   4
         Tag             =   "00-Enter Location Code"
         Top             =   2829
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin Threed.SSCheck chkManual 
         Height          =   255
         Left            =   5190
         TabIndex        =   227
         Top             =   3162
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Exclude from Update All  "
         ForeColor       =   -2147483640
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin INFOHR_Controls.DateLookup dlpAsOf 
         Height          =   285
         Left            =   1350
         TabIndex        =   8
         Tag             =   "40-As of Date"
         Top             =   3480
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   235
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   1875
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1335
         TabIndex        =   236
         Tag             =   "00-Specific Division Desired"
         Top             =   1875
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   6720
         TabIndex        =   237
         Tag             =   "EDPT-Category"
         Top             =   2193
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   7920
         TabIndex        =   269
         Top             =   4620
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   7995
         TabIndex        =   268
         Top             =   4380
         Width           =   1020
      End
      Begin VB.Label lblPT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   5190
         TabIndex        =   238
         Top             =   2238
         Width           =   630
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   234
         Top             =   3510
         Width           =   1245
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
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
         Left            =   30
         TabIndex        =   233
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   30
         TabIndex        =   232
         Top             =   2238
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   30
         TabIndex        =   231
         Top             =   2556
         Width           =   420
      End
      Begin VB.Label lblDtRange 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   30
         TabIndex        =   230
         Top             =   3192
         Width           =   1035
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   30
         TabIndex        =   229
         Top             =   2874
         Width           =   615
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         Left            =   5190
         TabIndex        =   226
         Top             =   2874
         Width           =   1260
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Update Method"
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
         Left            =   30
         TabIndex        =   225
         Top             =   3885
         Width           =   1110
      End
      Begin VB.Label lblSelCri 
         Caption         =   "Selection Criteria"
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
         TabIndex        =   18
         Top             =   1620
         Width           =   1575
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Entitlement"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   17
         Top             =   4620
         Width           =   960
      End
      Begin VB.Label lblHeading 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Service Ranges (in Months)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   4620
         Width           =   2370
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Entitlement Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   5190
         TabIndex        =   15
         Top             =   2556
         Width           =   1455
      End
      Begin VB.Label textMulti 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Union Code and FT/PT/SE/TR/OT will be validated from the Employee Basic Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   4350
         Visible         =   0   'False
         Width           =   7455
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status"
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
         Index           =   3
         Left            =   5190
         TabIndex        =   13
         Top             =   1920
         Width           =   1350
      End
   End
   Begin VB.Frame VacFram 
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   180
      TabIndex        =   19
      Top             =   4860
      Width           =   11235
      Begin Threed.SSFrame frmDH 
         Height          =   470
         Index           =   0
         Left            =   4890
         TabIndex        =   263
         Top             =   -75
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   811
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.64
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   0
            Left            =   1770
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   210
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   0
            Left            =   930
            TabIndex        =   24
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   210
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   23
            Tag             =   "Entitlement measured in days"
            Top             =   210
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   1
         Left            =   4890
         TabIndex        =   247
         Top             =   253
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   1
            Left            =   1770
            TabIndex        =   32
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   30
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   31
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   2
         Left            =   4890
         TabIndex        =   248
         Top             =   581
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   2
            Left            =   1770
            TabIndex        =   39
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   37
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   2
            Left            =   930
            TabIndex        =   38
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   3
         Left            =   4890
         TabIndex        =   249
         Top             =   909
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   3
            Left            =   1770
            TabIndex        =   46
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   3
            Left            =   930
            TabIndex        =   45
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   44
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   4
         Left            =   4890
         TabIndex        =   267
         Top             =   1237
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   4
            Left            =   1770
            TabIndex        =   53
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   51
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   4
            Left            =   930
            TabIndex        =   52
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   5
         Left            =   4890
         TabIndex        =   250
         Top             =   1565
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   5
            Left            =   1770
            TabIndex        =   60
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   58
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   5
            Left            =   930
            TabIndex        =   59
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   6
         Left            =   4890
         TabIndex        =   251
         Top             =   1893
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   6
            Left            =   1770
            TabIndex        =   67
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   6
            Left            =   930
            TabIndex        =   66
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   6
            Left            =   90
            TabIndex        =   65
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   7
         Left            =   4890
         TabIndex        =   252
         Top             =   2221
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   7
            Left            =   1770
            TabIndex        =   74
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   72
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   7
            Left            =   930
            TabIndex        =   73
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   8
         Left            =   4890
         TabIndex        =   253
         Top             =   2549
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   8
            Left            =   1770
            TabIndex        =   81
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   79
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   8
            Left            =   930
            TabIndex        =   80
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   9
         Left            =   4890
         TabIndex        =   254
         Top             =   2877
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   9
            Left            =   1770
            TabIndex        =   88
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   86
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   9
            Left            =   930
            TabIndex        =   87
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   10
         Left            =   4890
         TabIndex        =   255
         Top             =   3205
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.51
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   10
            Left            =   1770
            TabIndex        =   95
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   93
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   10
            Left            =   930
            TabIndex        =   94
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   31.73
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   11
         Left            =   4890
         TabIndex        =   195
         Top             =   3533
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   11
            Left            =   1770
            TabIndex        =   102
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   100
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   11
            Left            =   930
            TabIndex        =   101
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   12
         Left            =   4890
         TabIndex        =   196
         Top             =   3861
         Width           =   2625
         _Version        =   65536
         _ExtentX        =   4630
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   12
            Left            =   1770
            TabIndex        =   109
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   12
            Left            =   930
            TabIndex        =   108
            TabStop         =   0   'False
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   12
            Left            =   90
            TabIndex        =   107
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   810
            _Version        =   65536
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   13
         Left            =   4890
         TabIndex        =   256
         Top             =   4189
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   13
            Left            =   1770
            TabIndex        =   116
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   13
            Left            =   90
            TabIndex        =   114
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   13
            Left            =   930
            TabIndex        =   115
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   14
         Left            =   4890
         TabIndex        =   257
         Top             =   4517
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   14
            Left            =   1770
            TabIndex        =   123
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   121
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   14
            Left            =   930
            TabIndex        =   122
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   15
         Left            =   4890
         TabIndex        =   258
         Top             =   4845
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   15
            Left            =   1770
            TabIndex        =   130
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   15
            Left            =   90
            TabIndex        =   128
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   15
            Left            =   930
            TabIndex        =   129
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   16
         Left            =   4890
         TabIndex        =   259
         Top             =   5173
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   16
            Left            =   1770
            TabIndex        =   137
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   135
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   16
            Left            =   930
            TabIndex        =   136
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   17
         Left            =   4890
         TabIndex        =   260
         Top             =   5520
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   17
            Left            =   1770
            TabIndex        =   144
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   17
            Left            =   90
            TabIndex        =   142
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   17
            Left            =   930
            TabIndex        =   143
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   18
         Left            =   4890
         TabIndex        =   261
         Top             =   5829
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   18
            Left            =   1770
            TabIndex        =   151
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   18
            Left            =   90
            TabIndex        =   149
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   18
            Left            =   930
            TabIndex        =   150
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   19
         Left            =   4890
         TabIndex        =   262
         Top             =   6157
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   19
            Left            =   1770
            TabIndex        =   158
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   19
            Left            =   90
            TabIndex        =   156
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   19
            Left            =   930
            TabIndex        =   157
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   20
         Left            =   4890
         TabIndex        =   264
         Top             =   6485
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   20
            Left            =   1770
            TabIndex        =   165
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   20
            Left            =   90
            TabIndex        =   163
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   20
            Left            =   930
            TabIndex        =   164
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   21
         Left            =   4890
         TabIndex        =   265
         Top             =   6813
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   21
            Left            =   1770
            TabIndex        =   172
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   21
            Left            =   90
            TabIndex        =   170
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   21
            Left            =   930
            TabIndex        =   171
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   22
         Left            =   4890
         TabIndex        =   242
         Top             =   7141
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   22
            Left            =   1770
            TabIndex        =   179
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   22
            Left            =   90
            TabIndex        =   177
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   22
            Left            =   930
            TabIndex        =   178
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   23
         Left            =   4890
         TabIndex        =   244
         Top             =   7469
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   23
            Left            =   1770
            TabIndex        =   186
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   23
            Left            =   90
            TabIndex        =   184
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   23
            Left            =   930
            TabIndex        =   185
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Tag             =   "11-Service is greater than this number"
         Top             =   120
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   0
         Left            =   2115
         TabIndex        =   21
         Tag             =   "10-Service is less than this number"
         Top             =   120
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Tag             =   "11-Service is greater than this number"
         Top             =   448
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   1
         Left            =   2115
         TabIndex        =   28
         Tag             =   "10-Service is less than this number"
         Top             =   448
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   2
         Left            =   0
         TabIndex        =   34
         Tag             =   "11-Service is greater than this number"
         Top             =   776
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   2
         Left            =   2115
         TabIndex        =   35
         Tag             =   "10-Service is less than this number"
         Top             =   776
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   1
         Left            =   3690
         TabIndex        =   29
         Tag             =   "11-Entitlement Amount"
         Top             =   448
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   2
         Left            =   3690
         TabIndex        =   36
         Tag             =   "11-Entitlement Amount"
         Top             =   776
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   0
         Left            =   3690
         TabIndex        =   22
         Tag             =   "11-Entitlement Amount"
         Top             =   120
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   3
         Left            =   0
         TabIndex        =   41
         Tag             =   "11-Service is greater than this number"
         Top             =   1104
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   3
         Left            =   2115
         TabIndex        =   42
         Tag             =   "10-Service is less than this number"
         Top             =   1104
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   48
         Tag             =   "11-Service is greater than this number"
         Top             =   1432
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   4
         Left            =   2115
         TabIndex        =   49
         Tag             =   "10-Service is less than this number"
         Top             =   1432
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   5
         Left            =   0
         TabIndex        =   55
         Tag             =   "11-Service is greater than this number"
         Top             =   1760
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   5
         Left            =   2115
         TabIndex        =   56
         Tag             =   "10-Service is less than this number"
         Top             =   1760
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   4
         Left            =   3690
         TabIndex        =   50
         Tag             =   "11-Entitlement Amount"
         Top             =   1432
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   5
         Left            =   3690
         TabIndex        =   57
         Tag             =   "11-Entitlement Amount"
         Top             =   1760
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   6
         Left            =   0
         TabIndex        =   62
         Tag             =   "11-Service is greater than this number"
         Top             =   2088
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   6
         Left            =   2115
         TabIndex        =   63
         Tag             =   "10-Service is less than this number"
         Top             =   2088
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   7
         Left            =   0
         TabIndex        =   69
         Tag             =   "11-Service is greater than this number"
         Top             =   2416
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   7
         Left            =   2115
         TabIndex        =   70
         Tag             =   "10-Service is less than this number"
         Top             =   2416
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   8
         Left            =   0
         TabIndex        =   76
         Tag             =   "11-Service is greater than this number"
         Top             =   2744
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   8
         Left            =   2115
         TabIndex        =   77
         Tag             =   "10-Service is less than this number"
         Top             =   2744
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   7
         Left            =   3690
         TabIndex        =   71
         Tag             =   "11-Entitlement Amount"
         Top             =   2416
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   8
         Left            =   3690
         TabIndex        =   78
         Tag             =   "11-Entitlement Amount"
         Top             =   2744
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   6
         Left            =   3690
         TabIndex        =   64
         Tag             =   "11-Entitlement Amount"
         Top             =   2088
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   9
         Left            =   0
         TabIndex        =   83
         Tag             =   "11-Service is greater than this number"
         Top             =   3072
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   9
         Left            =   2115
         TabIndex        =   84
         Tag             =   "10-Service is less than this number"
         Top             =   3072
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   9
         Left            =   3690
         TabIndex        =   85
         Tag             =   "11-Entitlement Amount"
         Top             =   3072
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   90
         Tag             =   "11-Service is greater than this number"
         Top             =   3400
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   10
         Left            =   2115
         TabIndex        =   91
         Tag             =   "10-Service is less than this number"
         Top             =   3400
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   97
         Tag             =   "11-Service is greater than this number"
         Top             =   3728
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   11
         Left            =   2115
         TabIndex        =   98
         Tag             =   "10-Service is less than this number"
         Top             =   3728
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   10
         Left            =   3690
         TabIndex        =   92
         Tag             =   "11-Entitlement Amount"
         Top             =   3400
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   11
         Left            =   3690
         TabIndex        =   99
         Tag             =   "11-Entitlement Amount"
         Top             =   3728
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   12
         Left            =   0
         TabIndex        =   104
         Tag             =   "11-Service is greater than this number"
         Top             =   4056
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   12
         Left            =   2115
         TabIndex        =   105
         Tag             =   "10-Service is less than this number"
         Top             =   4056
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   13
         Left            =   0
         TabIndex        =   111
         Tag             =   "11-Service is greater than this number"
         Top             =   4384
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   13
         Left            =   2115
         TabIndex        =   112
         Tag             =   "10-Service is less than this number"
         Top             =   4384
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   14
         Left            =   0
         TabIndex        =   118
         Tag             =   "11-Service is greater than this number"
         Top             =   4712
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   14
         Left            =   2115
         TabIndex        =   119
         Tag             =   "10-Service is less than this number"
         Top             =   4712
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   12
         Left            =   3690
         TabIndex        =   106
         Tag             =   "11-Entitlement Amount"
         Top             =   4056
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   13
         Left            =   3690
         TabIndex        =   113
         Tag             =   "11-Entitlement Amount"
         Top             =   4384
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   14
         Left            =   3690
         TabIndex        =   120
         Tag             =   "11-Entitlement Amount"
         Top             =   4712
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   15
         Left            =   0
         TabIndex        =   125
         Tag             =   "11-Service is greater than this number"
         Top             =   5040
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   15
         Left            =   2115
         TabIndex        =   126
         Tag             =   "10-Service is less than this number"
         Top             =   5040
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   132
         Tag             =   "11-Service is greater than this number"
         Top             =   5368
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   16
         Left            =   2115
         TabIndex        =   133
         Tag             =   "10-Service is less than this number"
         Top             =   5368
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   15
         Left            =   3690
         TabIndex        =   127
         Tag             =   "11-Entitlement Amount"
         Top             =   5040
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   16
         Left            =   3690
         TabIndex        =   134
         Tag             =   "11-Entitlement Amount"
         Top             =   5368
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   17
         Left            =   0
         TabIndex        =   139
         Tag             =   "11-Service is greater than this number"
         Top             =   5696
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   17
         Left            =   2115
         TabIndex        =   140
         Tag             =   "10-Service is less than this number"
         Top             =   5696
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   17
         Left            =   3690
         TabIndex        =   141
         Tag             =   "11-Entitlement Amount"
         Top             =   5696
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   18
         Left            =   0
         TabIndex        =   146
         Tag             =   "11-Service is greater than this number"
         Top             =   6024
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   18
         Left            =   2115
         TabIndex        =   147
         Tag             =   "10-Service is less than this number"
         Top             =   6024
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   19
         Left            =   0
         TabIndex        =   153
         Tag             =   "11-Service is greater than this number"
         Top             =   6352
         Width           =   760
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   19
         Left            =   2115
         TabIndex        =   154
         Tag             =   "10-Service is less than this number"
         Top             =   6352
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   18
         Left            =   3690
         TabIndex        =   148
         Tag             =   "11-Entitlement Amount"
         Top             =   6024
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   19
         Left            =   3690
         TabIndex        =   155
         Tag             =   "11-Entitlement Amount"
         Top             =   6352
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   160
         Tag             =   "11-Service is greater than this number"
         Top             =   6680
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   20
         Left            =   2115
         TabIndex        =   161
         Tag             =   "10-Service is less than this number"
         Top             =   6680
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   20
         Left            =   3690
         TabIndex        =   162
         Tag             =   "11-Entitlement Amount"
         Top             =   6680
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   21
         Left            =   0
         TabIndex        =   167
         Tag             =   "11-Service is greater than this number"
         Top             =   7008
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   21
         Left            =   2115
         TabIndex        =   168
         Tag             =   "10-Service is less than this number"
         Top             =   7008
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   21
         Left            =   3690
         TabIndex        =   169
         Tag             =   "11-Entitlement Amount"
         Top             =   7008
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   22
         Left            =   0
         TabIndex        =   174
         Tag             =   "11-Service is greater than this number"
         Top             =   7336
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   22
         Left            =   2115
         TabIndex        =   175
         Tag             =   "10-Service is less than this number"
         Top             =   7336
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   22
         Left            =   3690
         TabIndex        =   176
         Tag             =   "11-Entitlement Amount"
         Top             =   7336
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   23
         Left            =   0
         TabIndex        =   181
         Tag             =   "11-Service is greater than this number"
         Top             =   7664
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   23
         Left            =   2115
         TabIndex        =   182
         Tag             =   "10-Service is less than this number"
         Top             =   7664
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medLTServ 
         Height          =   285
         Index           =   24
         Left            =   0
         TabIndex        =   188
         Tag             =   "11-Service is greater than this number"
         Top             =   8010
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medGTServ 
         Height          =   285
         Index           =   24
         Left            =   2115
         TabIndex        =   189
         Tag             =   "10-Service is less than this number"
         Top             =   8010
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   24
         Left            =   3690
         TabIndex        =   190
         Tag             =   "11-Entitlement Amount"
         Top             =   8010
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin Threed.SSFrame frmDH 
         Height          =   480
         Index           =   24
         Left            =   4890
         TabIndex        =   266
         Top             =   7815
         Width           =   2640
         _Version        =   65536
         _ExtentX        =   4657
         _ExtentY        =   847
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   16.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Font3D          =   1
         ShadowStyle     =   1
         Begin Threed.SSOption optD 
            Height          =   195
            Index           =   24
            Left            =   90
            TabIndex        =   191
            Tag             =   "Entitlement measured in days"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   1323
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Days"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
         End
         Begin Threed.SSOption optH 
            Height          =   195
            Index           =   24
            Left            =   930
            TabIndex        =   192
            Tag             =   "Entitlement measured in hours"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "Hours"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Threed.SSOption optF 
            Height          =   195
            Index           =   24
            Left            =   1770
            TabIndex        =   193
            TabStop         =   0   'False
            Tag             =   "Entitlement Measured in FTE#"
            Top             =   240
            Width           =   750
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   78
            Caption         =   "FTE#"
            ForeColor       =   -2147483640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   23
         Left            =   3690
         TabIndex        =   183
         Tag             =   "11-Entitlement Amount"
         Top             =   7664
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medEntitle 
         Height          =   285
         Index           =   3
         Left            =   3690
         TabIndex        =   43
         Tag             =   "11-Entitlement Amount"
         Top             =   1104
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0.0000"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   0
         Left            =   7800
         TabIndex        =   26
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   120
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   1
         Left            =   7800
         TabIndex        =   33
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   448
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   2
         Left            =   7800
         TabIndex        =   40
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   776
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   3
         Left            =   7800
         TabIndex        =   47
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   1104
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   4
         Left            =   7800
         TabIndex        =   54
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   1432
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   5
         Left            =   7800
         TabIndex        =   61
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   1760
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   6
         Left            =   7800
         TabIndex        =   68
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   2088
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   7
         Left            =   7800
         TabIndex        =   75
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   2416
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   8
         Left            =   7800
         TabIndex        =   82
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   2744
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   9
         Left            =   7800
         TabIndex        =   89
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   3072
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   10
         Left            =   7800
         TabIndex        =   96
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   3400
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   11
         Left            =   7800
         TabIndex        =   103
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   3728
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   12
         Left            =   7800
         TabIndex        =   110
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   4056
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   13
         Left            =   7800
         TabIndex        =   117
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   4384
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   14
         Left            =   7800
         TabIndex        =   124
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   4712
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   15
         Left            =   7800
         TabIndex        =   131
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   5040
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   16
         Left            =   7800
         TabIndex        =   138
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   5368
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   17
         Left            =   7800
         TabIndex        =   145
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   5696
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   18
         Left            =   7800
         TabIndex        =   152
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   6024
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   19
         Left            =   7800
         TabIndex        =   159
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   6352
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   20
         Left            =   7800
         TabIndex        =   166
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   6680
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   21
         Left            =   7800
         TabIndex        =   173
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   7008
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   22
         Left            =   7800
         TabIndex        =   180
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   7336
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   23
         Left            =   7800
         TabIndex        =   187
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   7664
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medMax 
         Height          =   285
         Index           =   24
         Left            =   7800
         TabIndex        =   194
         Tag             =   "10-Maximum Entitlement can be"
         Top             =   8010
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   ">    Service  "
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
         Index           =   24
         Left            =   930
         TabIndex        =   246
         Top             =   8055
         Width           =   900
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   23
         Left            =   840
         TabIndex        =   245
         Top             =   7709
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   22
         Left            =   840
         TabIndex        =   243
         Top             =   7381
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   21
         Left            =   840
         TabIndex        =   241
         Top             =   7053
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   240
         Top             =   6725
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   19
         Left            =   840
         TabIndex        =   223
         Top             =   6069
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   222
         Top             =   5741
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   17
         Left            =   840
         TabIndex        =   221
         Top             =   6397
         Width           =   1095
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   810
         TabIndex        =   213
         Top             =   5413
         Width           =   1095
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   212
         Top             =   2133
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   211
         Top             =   2461
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   210
         Top             =   2789
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   209
         Top             =   1149
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   4
         Left            =   840
         TabIndex        =   208
         Top             =   1477
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   3
         Left            =   840
         TabIndex        =   207
         Top             =   1805
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   2
         Left            =   840
         TabIndex        =   206
         Top             =   821
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   205
         Top             =   493
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   0
         Left            =   840
         TabIndex        =   204
         Top             =   165
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   203
         Top             =   4101
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Left            =   840
         TabIndex        =   202
         Top             =   4429
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   12
         Left            =   840
         TabIndex        =   201
         Top             =   4757
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   13
         Left            =   840
         TabIndex        =   200
         Top             =   3445
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   14
         Left            =   840
         TabIndex        =   199
         Top             =   3773
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   15
         Left            =   840
         TabIndex        =   198
         Top             =   3117
         Width           =   1080
      End
      Begin VB.Label lblSer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "<=  Service  =>"
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
         Index           =   16
         Left            =   840
         TabIndex        =   197
         Top             =   5085
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSHrsEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim snapEntitle As New ADODB.Recordset     'user vier
Dim fglbWDate$, fglbWDateS$
Dim NumAddRec%
Dim fglbSick%
Dim fglbVac%

Dim fglbNew As Boolean
Dim Actn
Dim AddChgDel As String

Dim fglbSDate As Variant
Dim fglbMaxRange%
Dim fglbCompMonthly%

Dim ffieldEntitle$    ' ED_VAC or ED_SICK for name of field for entitlement
Dim ffieldPEntitle$     ' ED_PVAC or ED_PSICK for previous entitlement's field name
Dim fglbCode$           ' are we dealing with Vac/Sick records?"
Dim fglbMaxRanges%
Dim glbFrmCaption$, glbErrNum&

Dim ControlsShown As Boolean
Dim ODIV, ODept, oOrg, oFDate, OTDate, oEMP, oEmpMode, oHETYPE, oAsOf
Dim OLoc, OSection
Dim OManual

Dim FlagRefresh As Boolean

Dim SnapAddEntitle As New ADODB.Recordset
Dim fglbESQLQ, fglbWSQLQ, fglbVSQLQ
Dim orgEffDate

Private Function chkMUEntitle(Optional xOKClick)
Dim x%, Y%

chkMUEntitle = False

On Error GoTo chkMUEntitle_Err
For x% = 0 To 4
If Len(clpCode(x%).Text) > 0 And clpCode(x%).Caption = "Unassigned" Then
    MsgBox "If Code entered it must be known"
    clpCode(x%).SetFocus
    Exit Function
End If
Next x%

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If
If Len(clpDiv.Text) < 1 Then
    If glbDIVCount = 1 And glbLinamar Then
        MsgBox lStr("Division is required field")
         clpDiv.SetFocus
        Exit Function
    End If
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("If Division Entered - it must be known")
         clpDiv.SetFocus
        Exit Function
    End If
End If
If Not glbCBrant Then 'Ticket #12524
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid From Date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Not glbCBrant Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid To Date"
        dlpTo.SetFocus
        Exit Function
    End If
End If

If Not IsDate(dlpAsOf.Text) Then
    MsgBox "Invalid Effective Date"
    dlpAsOf.SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) < 1 Then
    MsgBox "Entitlement Code is required field"
    clpCode(2).SetFocus
    Exit Function
Else
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox "If Code Entered - it must be known"
        clpCode(2).SetFocus
        Exit Function
    End If
End If

'If FLEX Logic - Cannot create Entitlement logic with '-' suffixed Entitlement Code
If Len(clpCode(2).Text) > 0 Then
    If Right(clpCode(2).Text, 1) = "-" Then
        MsgBox "Invalid Entitlement code. The Entitlement Code cannot have '-' suffixed to it."
        clpCode(2).SetFocus
        Exit Function
    End If
End If

'7.9 Enhancement - Cannot create OTs or CTs code entitlements
If Len(clpCode(2).Text) > 0 Then
    If Left(clpCode(2).Text, 2) = "OT" Or Left(clpCode(2).Text, 2) = "CT" Then
        MsgBox "Invalid Entitlement code. The Hourly Entitlement cannot be set for codes with 'OT' or 'CT' prefix to it."
        clpCode(2).SetFocus
        Exit Function
    End If
End If

'7.9 Enhancement - Warn to not create VACs or SICs code entitlements. This is for those client using ESS/TS.
If Len(clpCode(2).Text) > 0 Then
    If Left(clpCode(2).Text, 3) = "VAC" Or Left(clpCode(2).Text, 3) = "SIC" Then
        MsgBox "Please avoid creating Hourly Entitlements for codes prefixed 'VAC' and 'SIC', if using ESS/Timesheet Web Modules.", vbExclamation, "info:HR - Hourly Entitlement"
        'clpCode(2).SetFocus
        'Exit Function
    End If
End If

If glbWFC Then
    If Len(clpCode(3).Text) = 0 Then
        MsgBox lStr("Section is required field")
        clpCode(3).SetFocus
        Exit Function
    End If
End If
If Len(medLTServ(0)) < 1 Then
    MsgBox "You must have at least one Service Range Entry."
    If medLTServ(0).Enabled Then medLTServ(0).SetFocus
    Exit Function
End If

fglbMaxRanges% = 0  ' 0 is first range

Dim intRangesSet%
intRangesSet% = 0    ' 1 to 4 with 0 implying none
If Len(medLTServ(19)) = 0 Then
    medGTServ(19) = ""
Else
    If medLTServ(19) = 0 Then
        medLTServ(19) = ""
        medGTServ(19) = ""
    End If
End If


For x% = 0 To 24
    If Len(medLTServ(x%)) > 0 Then
        If Not IsNumeric(medLTServ(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medLTServ(x%).SetFocus
            Exit Function
        End If
    End If
    If Len(medGTServ(x%)) > 0 Then
        If Not IsNumeric(medGTServ(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medGTServ(x%).SetFocus
            Exit Function
        End If
    End If
    If Len(medEntitle(x%)) > 0 Then
        If Not IsNumeric(medEntitle(x%)) Then
            MsgBox "Data Entered Must Be Numeric"
            medEntitle(x%).SetFocus
            Exit Function
        End If
    Else
        If Len(medLTServ(x%)) > 0 Or Len(medGTServ(x%)) > 0 Then
             MsgBox "Numeric Value For Entitlement Must Be Entered"
             medEntitle(x%).SetFocus
             Exit Function
        End If
    End If


    If Len(medLTServ(x%)) < 1 And Len(medGTServ(x%)) > 1 Then  ' missed one
        MsgBox "Ranges must be sequential"
        medLTServ(x%).SetFocus
        Exit Function
    End If
    If Len(medGTServ(x%)) > 0 Then
        If glbFrench Then
            If CDbl(medLTServ(x%)) > CDbl(medGTServ(x%)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        Else
            If Val(medLTServ(x%)) > Val(medGTServ(x%)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        End If
    End If
    If x% > 0 And Len(medLTServ(x%)) > 0 Then
        If glbFrench Then
            If CDbl(medLTServ(x%)) < CDbl(medGTServ(x% - 1)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        Else
            If Val(medLTServ(x%)) < Val(medGTServ(x% - 1)) Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        End If
    End If
    If x% > 0 And Len(medGTServ(x%)) > 0 Then
        If glbFrench Then
            If CDbl(medGTServ(x%)) < CDbl(medGTServ(x% - 1)) And CDbl(medGTServ(x%)) <> 0 Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        Else
            If Val(medGTServ(x%)) < Val(medGTServ(x% - 1)) And Val(medGTServ(x%)) <> 0 Then
                MsgBox "Ranges must be sequential"
                medLTServ(x%).SetFocus
                Exit Function
            End If
        End If
    End If
    If Len(medLTServ(x%)) < 1 Then Exit For  ' missed one
    intRangesSet% = intRangesSet% + 1
Next x%

If intRangesSet% = 0 Then
    MsgBox "At least one Service level must be set"
    medLTServ(0).SetFocus
    Exit Function
End If

If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
    For x% = 0 To 24
        If Len(medMax(x%)) < 1 Then
            medMax(x%) = 0
        End If
    Next x%
End If

'Ticket #29617 - Mississaugas of Scugog Island First Nation - Only allow Hours
If glbCompSerial = "S/N - 2485W" Then
    If optH(0) = False Then
        MsgBox "Only Entitlement in Hours allowed"
        'optH(0).SetFocus
        Exit Function
    End If
End If

If IsMissing(xOKClick) Then
    If orgEffDate <> dlpAsOf.Text Then
        MsgBox "Effective Date has been changed. Please Save the changes before doing the Update."
        Exit Function
    End If
End If

chkMUEntitle = True

Exit Function

chkMUEntitle_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub cmdAddEnt_Click()
'**************** Hourly Entitlement Master Add Procedure
Dim SQLQ As String, Msg$, x%
Dim Title$, DgDef As Variant, Response%
On Error GoTo AddN_Err
If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
'
Title$ = "Mass Hourly Entitlement Records Addition"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If
'
AddChgDel = "A"

If Not chkMUEntitle() Then Exit Sub

If Not modInsSelection() Then Exit Sub   'laura 03/04/98

Call EntReCalcHr  'laura dec 15, 1997

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

data1.Refresh

Call Display_Value

Screen.MousePointer = DEFAULT

MsgBox "Records Added Successfully"

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HOURLY ENTITLEMENTS", "Add")
Resume Next

End Sub

Public Sub cmdCancel_Click()
fglbNew = False
data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim SQLQ, Msg, a%
If data1.Recordset.BOF And data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "The Hourly Entitlement Rules?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call getWSQLQ("C")

SQLQ = "DELETE FROM HR_HOURLYENT WHERE " & fglbVSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

End Sub

Private Sub Accum_Click(Value As Integer)
    If Accum.Value = True Then
        txtUpdMethod.Text = "A"
    ElseIf Replace.Value = True Then
        txtUpdMethod.Text = "R"
    End If
End Sub

Private Sub cmdDeleteAll_Click()
Dim a As Integer
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

On Error GoTo DelAll_Err

If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If orgEffDate <> dlpAsOf.Text Then
    MsgBox "Effective Date has been changed. Please Save the changes before doing the Delete."
    Exit Sub
End If

Title$ = "Mass Hourly Entitlement Records Delete ALL"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL the Hourly Entitlement records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

AddChgDel = "D"

If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.MoveFirst
    Do
        Call Display_Value
        
        If chkManual.Value = False Then
            If chkMUEntitle() Then
            
                recCount = getRecordCount_Delete(False)
                If recCount > 0 Then
                    Msg$ = Str(recCount)
                    If recCount = 1 Then Msg$ = Msg$ & " Hourly Entitlement Record " Else Msg$ = Msg$ & " Hourly Entitlement Records "
                    Msg$ = Msg$ & "will be Deleted for this group. " & vbCrLf & vbCrLf & "Do you want to proceed?"
                    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
                    If Response = IDNO Then
                        GoTo Next_Hourly
                    End If
                Else
                    MsgBox "No Hourly Entitlement record found to delete for this group."
                    GoTo Next_Hourly
                End If
                
                x% = modDelRecs()
            End If
        End If
Next_Hourly:
        data1.Recordset.MoveNext
    Loop Until data1.Recordset.EOF
End If

Call EntReCalcHr  'laura dec 15, 1997

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

Screen.MousePointer = DEFAULT

MsgBox "All Records Deleted Successfully"

Exit Sub

DelAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDeleteAll", "Hourly Entitlement", "Delete All")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdDeleteEnt_Click()
Dim a As Integer
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

AddChgDel = "D"

If Not chkMUEntitle() Then Exit Sub

Title$ = "Mass Hourly Entitlement Records Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete(False)
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Hourly Entitlement Record " Else Msg$ = Msg$ & " Hourly Entitlement Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Hourly Entitlement record found to delete."
    Exit Sub
End If

x% = modDelRecs()

Call EntReCalcHr  'laura dec 15, 1997

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

Screen.MousePointer = DEFAULT

MsgBox "Records Deleted Successfully."

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "ATTEND", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdModify_Click()

ODIV = clpDiv.Text
ODept = clpDept.Text
oOrg = clpCode(0).Text
oFDate = dlpFrom.Text
OTDate = dlpTo.Text
oAsOf = dlpAsOf.Text
oEMP = clpCode(1).Text
oEmpMode = clpPT.Text
oHETYPE = clpCode(2).Text
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    OLoc = clpProv.Text
Else
    OLoc = clpCode(4).Text
End If
OSection = clpCode(3).Text

orgEffDate = dlpAsOf.Text
OManual = chkManual.Value

Actn = "M"
End Sub

Public Sub cmdNew_Click()
Dim x

For x = 0 To 24
    medLTServ(x) = ""
    medGTServ(x) = ""
    medEntitle(x) = ""
    optD(x) = True
    optH(x) = False
    optF(x) = False
    If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
        medMax(x) = ""
    End If
Next

clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
dlpFrom.Text = ""
dlpTo.Text = ""
dlpAsOf.Text = ""
clpCode(1).Text = ""
clpCode(2).Text = ""
clpCode(3).Text = ""

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    clpProv.Text = ""
Else
    clpCode(4).Text = ""
End If

clpPT.Text = ""

Actn = "A"

fglbNew = True

SET_UP_MODE

clpDiv.SetFocus

End Sub

Public Sub cmdOK_Click()
Dim x%, Y%, xUnion, xPT, SQLQ, SQLQW
Dim xStr
Dim rsVE As New ADODB.Recordset
Dim rsVT As New ADODB.Recordset
Dim glbiOneWhere As Boolean
Dim bmk As Variant
Dim xFromDate As Date
Dim xToDate As Date
Dim xType As String


On Error GoTo AddN_Err

If data1.Recordset.EOF And data1.Recordset.BOF Then
    bmk = 0 'Ticket #11885 Frank Oct 11th, 2006
Else
    bmk = data1.Recordset.Bookmark
End If

If Not chkMUEntitle("OKClick") Then Exit Sub

For x% = 0 To 24
    If Not IsNumeric(medLTServ(x%)) Then Exit For
    If Not IsNumeric(medGTServ(x%)) Then
        medGTServ(x%) = 0
    Else
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If medLTServ(x%) > 0 And medGTServ(x%) = 0 Then medGTServ(x%) = 9999999
Next

If Actn = "M" Then
    Call getWSQLQ("O")
    SQLQ = "DELETE FROM HR_HOURLYENT WHERE " & fglbVSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HR_HOURLYENT WHERE " & fglbVSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You can not add duplicate record"
         clpDiv.SetFocus
        Exit Sub
    End If
End If

gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HR_HOURLYENT"
rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

For x% = 0 To 24
    If Len(medLTServ(x%)) > 0 Then
        rsVE.AddNew
        rsVE("EH_ORDER") = x + 1
        rsVE("EH_ORG_TABL") = "EDOR"
        rsVE("EH_ORG") = clpCode(0).Text
        rsVE("EH_PT") = clpPT.Text
        rsVE("EH_DIV") = clpDiv.Text
        rsVE("EH_DEPT") = clpDept.Text
        rsVE("EH_EMP_TABL") = "EDEM"
        rsVE("EH_EMP") = clpCode(1).Text
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
            rsVE("EH_LOC") = clpProv.Text
        Else
            rsVE("EH_LOC") = clpCode(4).Text
        End If
        rsVE("EH_SECTION") = clpCode(3).Text
        
        If Len(dlpFrom.Text) > 0 Then rsVE("EH_FDATE") = dlpFrom.Text
        If Len(dlpTo.Text) > 0 Then rsVE("EH_TDATE") = dlpTo.Text
        If Len(dlpAsOf.Text) > 0 Then rsVE("EH_EDATE") = dlpAsOf.Text
        
        rsVE("EH_HETYPE_TABL") = "ADRE"
        rsVE("EH_HETYPE") = clpCode(2).Text
        If glbFrench Then
            rsVE("EH_BMONTH") = Replace(medLTServ(x%), ",", ".")
        Else
            rsVE("EH_BMONTH") = medLTServ(x%)
        End If
        If glbFrench Then
            rsVE("EH_EMONTH") = Replace(medGTServ(x%), ",", ".")
        Else
            rsVE("EH_EMONTH") = medGTServ(x%)
        End If
        If glbFrench Then
            rsVE("EH_ENTITLE") = Replace(medEntitle(x%), ",", ".")
        Else
            rsVE("EH_ENTITLE") = medEntitle(x%)
        End If
        If optD(x%) Then rsVE("EH_TYPE") = "D"
        If optH(x%) Then rsVE("EH_TYPE") = "H"
        If optF(x%) Then rsVE("EH_TYPE") = "F"
        rsVE("EH_MANUAL") = chkManual.Value
        If Len(txtUpdMethod.Text) = 0 Then
            If Accum.Value = True Then rsVE("EH_UPDMETHOD") = "A"
            If Replace.Value Then rsVE("EH_UPDMETHOD") = "R"
        Else
            rsVE("EH_UPDMETHOD") = txtUpdMethod.Text
        End If
        If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
            rsVE("EH_MAX") = medMax(x%)
        End If
        rsVE.Update
    End If
Next
xFromDate = rsVE("EH_FDATE")
xToDate = rsVE("EH_TDATE")
xType = rsVE("EH_HETYPE")
xPT = rsVE("EH_PT")

rsVE.Close
gdbAdoIhr001.CommitTrans
data1.Refresh

If Not bmk = 0 Then
    data1.Recordset.Bookmark = bmk
End If

'data1.Recordset.Find "EH_FDATE=" & Format(xFromDate, "mm-dd-yyyy") '& " AND EH_TDATE = " & xToDate & " AND EH_HETYPE = '" & xType & "'"
'data1.Recordset.Find "EH_FDATE=" & CVDate(xFromDate) 'Ticket #27205 Franks 06/18/2015 for date format dd/mm/yyyy
data1.Recordset.MoveFirst                               'Ticket #29710 - Added this as the find below was giving blank row at times. The Find seems to work as forward find only.
data1.Recordset.Find "EH_FDATE=" & Date_SQL(xFromDate)  'Ticket #29710 - the above CVDATE() was giving error 'data type mismatch' for date format mmm dd/yy
data1.Recordset.Find "EH_HETYPE = '" & xType & "'"
'data1.Recordset.Find "EH_PT = '" & xPT & "'"

fglbNew = False

Call Display_Value

orgEffDate = dlpAsOf.Text

vbxTrueGrid.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

If Err.Number = -2147217887 Then '01/01/1200 can cause this error Ticket #18227
    MsgBox "    Invalid Date!    "
    gdbAdoIhr001.RollbackTrans
    Exit Sub
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "HOURLY ENTITLEMENTS", "UPDATE")
    Unload Me
End If

End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%

'mdPrint.Enabled = False
cmdPrintAll.Enabled = False

Me.vbxCrystal.Reset

Me.vbxCrystal.WindowTitle = "Hourly Entitlement Master Report"

Call setRptLabel(Me, 0) '1)

Me.vbxCrystal.Connect = RptODBC_SQL
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsent.rpt"

SQLQ = ""
SQLQ = SQLQ & "{HR_HOURLYENT.EH_DIV} = '" & clpDiv.Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_DEPT} = '" & clpDept.Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_ORG} = '" & clpCode(0).Text & "'"
If Len(dlpFrom.Text) > 0 Then
    dtYYY% = Year(dlpFrom.Text)
    dtMM% = month(dlpFrom.Text)
    dtDD% = Day(dlpFrom.Text)
    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_FDATE} in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(dlpTo.Text) > 0 Then
    dtYYY% = Year(dlpTo.Text)
    dtMM% = month(dlpTo.Text)
    dtDD% = Day(dlpTo.Text)
    SQLQ = SQLQ & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_EMP} = '" & clpCode(1).Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_PT} = '" & clpPT.Text & "' "
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_HETYPE} = '" & clpCode(2).Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_SECTION} = '" & clpCode(3).Text & "'"
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpProv.Text & "'"
Else
    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpCode(4).Text & "'"
End If
Me.vbxCrystal.SelectionFormula = SQLQ

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True
cmdPrintAll.Enabled = True
Call SET_UP_MODE
End Sub

Public Sub cmdView_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%

'cmdPrint.Enabled = False
cmdPrintAll.Enabled = False

Me.vbxCrystal.Reset

Me.vbxCrystal.WindowTitle = "Hourly Entitlement Master Report"

Call setRptLabel(Me, 0) '1)

Me.vbxCrystal.Connect = RptODBC_SQL
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsent.rpt"

SQLQ = ""
SQLQ = SQLQ & "{HR_HOURLYENT.EH_DIV} = '" & clpDiv.Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_DEPT} = '" & clpDept.Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_ORG} = '" & clpCode(0).Text & "'"
If Len(dlpFrom.Text) > 0 Then
    dtYYY% = Year(dlpFrom.Text)
    dtMM% = month(dlpFrom.Text)
    dtDD% = Day(dlpFrom.Text)
    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_FDATE} in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
If Len(dlpTo.Text) > 0 Then
    dtYYY% = Year(dlpTo.Text)
    dtMM% = month(dlpTo.Text)
    dtDD% = Day(dlpTo.Text)
    SQLQ = SQLQ & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
End If
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_EMP} = '" & clpCode(1).Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_PT} = '" & clpPT.Text & "' "
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_HETYPE} = '" & clpCode(2).Text & "'"
SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_SECTION} = '" & clpCode(3).Text & "'"
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpProv.Text & "'"
Else
    SQLQ = SQLQ & " AND {HR_HOURLYENT.EH_LOC} = '" & clpCode(4).Text & "'"
End If
Me.vbxCrystal.SelectionFormula = SQLQ

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

Call SET_UP_MODE
'cmdPrint.Enabled = True
cmdPrintAll.Enabled = True
End Sub

Private Sub cmdDeletePrv_Click()
Dim a As Integer
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

AddChgDel = "D"

If Not chkMUEntitle() Then Exit Sub

Title$ = "Mass Hourly Entitlement Records Delete Previous Year"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete(True)
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Previous Year Hourly Entitlement Record " Else Msg$ = Msg$ & " Previous Year Hourly Entitlement Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Previous Year Hourly Entitlement record found to delete."
    Exit Sub
End If

x% = modDelRecs(True)

'Call EntReCalcHr  'laura dec 15, 1997

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

Screen.MousePointer = DEFAULT

MsgBox "Records Deleted Successfully."

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDeletePrv", "HRENTHRS ", "Delete Prv Year")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdPrintAll_Click()
Dim RHeading As String, xReport, x%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%

cmdPrintAll.Enabled = False
'cmdPrint.Enabled = False
Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Hourly Entitlement Master Report"

Call setRptLabel(Me, 0) '1)

Me.vbxCrystal.Connect = RptODBC_SQL
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rghrsent.rpt"
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True
cmdPrintAll.Enabled = True
End Sub

Public Sub cmdUpdate_Click()
Dim SQLQ As String, Msg$, x%
Dim Title$, DgDef As Variant, Response%
Dim sFlag As Boolean
Dim recCount As Integer

On Error GoTo AddN_Err

If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

Title$ = "Mass Hourly Entitlement Records Update"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2
Msg$ = "Are you sure you want to Update Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)
If Response% = IDNO Then
    Exit Sub
End If

If Not chkMUEntitle() Then Exit Sub

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Hourly Entitlement Record " Else Msg$ = Msg$ & " Hourly Entitlement Records "
    Msg$ = Msg$ & "will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Hourly Entitlement record found to update."
    Exit Sub
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
 
sFlag = DoWork

data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

Screen.MousePointer = DEFAULT

If sFlag Then
    MsgBox "Records Updated Successfully."
End If

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HOURLY ENTITLEMENTS", "Add")
Resume Next
'ZAHOOR BUTT 01/11/2006

End Sub

Private Function DoWork() As Boolean
Dim sFlag As Boolean

Screen.MousePointer = DEFAULT

DoWork = False

If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Function
End If

AddChgDel = "A"

If glbCBrant Then
    If Not modInsSelectionCBrant() Then Exit Function
Else
    If Not modInsSelection() Then Exit Function
    'If Not modUpdateSelection() Then Exit Function
End If

Call EntReCalcHr

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

DoWork = True

End Function

Private Function CR_SnapAddEntitle()

Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$, strTm$, x%
Dim Dt As Variant

CR_SnapAddEntitle = False

On Error GoTo CR_SnapAddEntitle_Err

strTm$ = Time$

Dt = Date$

Call getWSQLQ("")

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT JH_DHRS,JH_FTENUM,ED_EMPNBR,ED_DHRS,ED_DOH,ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1 "
If glbOracle Then
    SQLQ = SQLQ & "FROM HREMP, HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT<>0"
Else
    SQLQ = SQLQ & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
End If

SQLQ = SQLQ & " AND " & fglbESQLQ

If SnapAddEntitle.State <> 0 Then SnapAddEntitle.Close
SnapAddEntitle.Open SQLQ, gdbAdoIhr001, adOpenStatic

NumAddRec% = SnapAddEntitle.RecordCount
Screen.MousePointer = DEFAULT
CR_SnapAddEntitle = True

Exit Function

CR_SnapAddEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapAddEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function CR_SnapEntitle()

Dim SQLQ As String

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

Call getWSQLQ("")

SQLQ = "SELECT * FROM qry_MU_Hourly "
SQLQ = SQLQ & " Where " & fglbESQLQ & " AND " & fglbWSQLQ

If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenStatic
Screen.MousePointer = DEFAULT
CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmdUpdateAll_Click()
Dim failed As String
Dim c As Long
Dim recCount As Integer
Dim Msg$
Dim Title$, DgDef As Variant, Response%

On Error GoTo Mod_Err

If Not gSec_Upd_Hrly_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If orgEffDate <> dlpAsOf.Text Then
    MsgBox "Effective Date has been changed. Please Save the changes before doing the Update."
    Exit Sub
End If

Title$ = "Mass Hourly Entitlement Records Update"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2
Msg$ = "Are you sure you want to Update ALL Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)
If Response% = IDNO Then
    Exit Sub
End If

failed = ""
c = 1
If data1.Recordset.EOF = False And data1.Recordset.BOF = False Then
    data1.Recordset.MoveFirst
    Do
        Call Display_Value
        
        If chkManual.Value = False Then
            If chkMUEntitle() Then
                recCount = getRecordCount_Add
                If recCount > 0 Then
                    Msg$ = Str(recCount)
                    If recCount = 1 Then Msg$ = Msg$ & " Hourly Entitlement Record " Else Msg$ = Msg$ & " Hourly Entitlement Records "
                    Msg$ = Msg$ & "will be Updated for this group. " & vbCrLf & vbCrLf & "Do you want to proceed?"
                    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)     ' Get user response.
                    If Response = IDNO Then
                        GoTo Next_Hourly
                    End If
                Else
                    MsgBox "No Hourly Entitlement record found to update for this group."
                    GoTo Next_Hourly
                End If
            
               If DoWork = False Then
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
                    If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
                    If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
                    If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
                    If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
                    If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
                    If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
                    If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
                    If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
               End If
            Else
                failed = failed & "Rule " & CStr(c) & ": "
                If Not IsNull(data1.Recordset("EH_DIV")) Then failed = failed & data1.Recordset("EH_DIV") & ", "
                If Not IsNull(data1.Recordset("EH_DEPT")) Then failed = failed & data1.Recordset("EH_DEPT") & ", "
                If Not IsNull(data1.Recordset("EH_ORG")) Then failed = failed & data1.Recordset("EH_ORG") & ", "
                If Not IsNull(data1.Recordset("EH_EMP")) Then failed = failed & data1.Recordset("EH_EMP") & ", "
                If Not IsNull(data1.Recordset("EH_PT")) Then failed = failed & data1.Recordset("EH_PT") & ", "
                If Not IsNull(data1.Recordset("EH_HETYPE")) Then failed = failed & data1.Recordset("EH_HETYPE") & ", "
                If Not IsNull(data1.Recordset("EH_FDATE")) Then failed = failed & data1.Recordset("EH_FDATE") & ", "
                If Not IsNull(data1.Recordset("EH_TDATE")) Then failed = failed & data1.Recordset("EH_TDATE") & ", "
                If Not IsNull(data1.Recordset("EH_EDATE")) Then failed = failed & data1.Recordset("EH_EDATE") & ", "
                If Not IsNull(data1.Recordset("EH_LOC")) Then failed = failed & data1.Recordset("EH_LOC") & ", "
                If Not IsNull(data1.Recordset("EH_SECTION")) Then failed = failed & data1.Recordset("EH_SECTION") & ", "
                failed = Left(failed, Len(failed) - 2) & vbCrLf
            End If
        End If
        
Next_Hourly:
        c = c + 1
        data1.Recordset.MoveNext
    Loop Until data1.Recordset.EOF
End If

data1.Refresh

Call Display_Value

orgEffDate = dlpAsOf.Text

Screen.MousePointer = DEFAULT

If Len(failed) = 0 Then
    MsgBox "All Rules applied", vbInformation + vbOKOnly, "Hourly Entitlements"
Else
    MsgBox "The Following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Hourly Entitlements"
End If

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Hourly", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Call INI_Controls(Me)
glbOnTop = "FRMSHRSENT"

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim x%
Dim SQLQ

glbOnTop = "FRMSHRSENT"

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    clpCode(4).Visible = False
    clpProv.Left = clpCode(4).Left
    clpProv.Top = clpCode(4).Top
    clpProv.Visible = True
    lblLocation.Caption = "Province"
    vbxTrueGrid.Columns(8).Caption = "Province"
End If

'Ticket #29617 - Mississaugas of Scugog Island First Nation
'Ticket #27729 Franks 03/14/2016 Carizon
If glbCompSerial = "S/N - 2430W" Or glbCompSerial = "S/N - 2485W" Then
    Call ScreenSetup(True)
Else
    Call ScreenSetup(False)
End If


FlagRefresh = False

data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT DISTINCT EH_DIV,EH_DEPT,EH_ORG,EH_FDATE,EH_TDATE,EH_EMP,EH_SECTION,EH_LOC,EH_PT,EH_HETYPE,EH_MANUAL,EH_EDATE,EH_UPDMETHOD FROM HR_HOURLYENT "
If glbDIVCount = 1 And glbLinamar Then
    SQLQ = SQLQ & " WHERE EH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
End If
If glbWFC Then 'Ticket #28553 Franks 05/03/2016
    SQLQ = SQLQ & " WHERE " & getWFCPlantSecurity("EH_SECTION")
End If

data1.RecordSource = SQLQ
data1.Refresh

ODIV = ""
ODept = ""
oOrg = ""
oFDate = ""
OTDate = ""
oAsOf = ""
oEMP = ""
oEmpMode = ""
oHETYPE = ""
OLoc = ""
OSection = ""
orgEffDate = ""
OManual = False

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select
Select Case glbEntOutStandingS$
    Case "2": fglbWDateS$ = "ED_DOH"
    Case "3": fglbWDateS$ = "ED_SENDTE"
    Case "4": fglbWDateS$ = "ED_LTHIRE"
    Case "5": fglbWDateS$ = "ED_USRDAT1"
    Case "6": fglbWDateS$ = "ED_UNION"
End Select

If UCase(glbCompEntVac$) = "M" Then
    vbxTrueGrid.Columns(3).Visible = False
End If
If glbWFC Then
    lblSection.FontBold = True
End If

Screen.MousePointer = HOURGLASS
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)

Call setRptCaption(Me)

If glbCBrant Then
    lblDtRange.Visible = False
    dlpFrom.Visible = False
    dlpTo.Visible = False
End If

Screen.MousePointer = DEFAULT

If glbMulti Then textMulti.Visible = True
textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

ST_UPD_MODE (False)

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
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
If Me.Height >= 6960 + VacFram.Height + panControls.Height + 230 Then
    scrControl.Value = 0
    VacFram.Top = 4570 '3960
    scrControl.Visible = False
    Exit Sub
End If
scrControl.Visible = True
scrControl.Max = VacFram.Height + panControls.Height + 6960 - Me.Height
scrControl.Left = Me.Width - scrControl.Width - 260
If Me.Height - scrControl.Top - panControls.Height - 400 > 0 Then
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 400
Else
    scrControl.Height = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."

Set frmUEntitle = Nothing  'carmen apr 2000
End Sub

Private Sub medEntitle_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medGTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medLTServ_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optD_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optD_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optF_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optF_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optH_Click(Index As Integer, Value As Integer)
    Call ST_OPT_VALUE
End Sub

Private Sub optH_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Replace_Click(Value As Integer)
    If Accum.Value = True Then
        txtUpdMethod.Text = "A"
    ElseIf Replace.Value = True Then
        txtUpdMethod.Text = "R"
    End If
End Sub

Private Sub scrControl_Change()
VacFram.Top = 4800 - scrControl.Value '4300 '3960 ' 4800
End Sub

Sub ST_UPD_MODE(TF As Boolean)
Dim x, FT
FT = Not TF
For x = 0 To 24
    medLTServ(x).Enabled = TF
    medGTServ(x).Enabled = TF
    medEntitle(x).Enabled = TF
    If x = 0 Then
        optD(x).Enabled = TF
        optH(x).Enabled = TF
        optF(x).Enabled = TF
    Else
        optD(x).Enabled = False
        optH(x).Enabled = False
        optF(x).Enabled = False
    End If
Next
clpDiv.Enabled = TF
clpDept.Enabled = TF
clpCode(0).Enabled = TF
dlpFrom.Enabled = TF
dlpTo.Enabled = TF
dlpAsOf.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    clpProv.Enabled = TF
Else
    clpCode(4).Enabled = TF
End If
clpPT.Enabled = TF
'cmdClose.Enabled = FT

If data1.Recordset.EOF And data1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
Else
'    cmdModify.Enabled = FT
'    cmdDelete.Enabled = FT
End If

ODIV = clpDiv.Text

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
'cmdNew.Enabled = FT
'cmdPrint.Enabled = FT
'cmdPrintAll.Enabled = FT
'cmdUpdate.Enabled = FT
'cmdAddEnt.Enabled = FT
'cmdDeleteEnt.Enabled = FT

'vbxTrueGrid.Enabled = FT

'Ticket #29617 - Mississaugas of Scugog Island First Nation
If glbCompSerial = "S/N - 2485W" Then
    optD(0).Enabled = False
    'optH(0).Enabled = False
    optF(0).Enabled = False
End If

Call modSetFGlobals("Vac")

End Sub

Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim x

For x = 0 To 24
    medLTServ(x) = ""
    medGTServ(x) = ""
    medEntitle(x) = ""
    optD(x) = True
    optH(x) = False
    optF(x) = False
    If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
        medMax(x) = ""
    End If
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
dlpFrom.Text = ""
dlpTo.Text = ""
dlpAsOf.Text = ""
clpCode(1).Text = ""
clpCode(2).Text = ""
clpCode(3).Text = ""

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    clpProv.Text = ""
Else
    clpCode(4).Text = ""
End If

clpPT.Text = ""

If Not data1.Recordset.EOF Then
    Call getWSQLQ("D")
    
    SQLQ = "SELECT * FROM HR_HOURLYENT WHERE " & fglbVSQLQ
    SQLQ = SQLQ & "Order By EH_DIV,EH_DEPT,EH_ORG, EH_FDATE,EH_EMP,EH_PT,EH_LOC,EH_SECTION,EH_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    
    If Not IsNull(data1.Recordset("EH_DIV")) Then clpDiv.Text = data1.Recordset("EH_DIV")
    If Not IsNull(data1.Recordset("EH_DEPT")) Then clpDept.Text = data1.Recordset("EH_DEPT")
    If Not IsNull(data1.Recordset("EH_ORG")) Then clpCode(0).Text = data1.Recordset("EH_ORG")
    If Not IsNull(data1.Recordset("EH_FDATE")) Then dlpFrom.Text = data1.Recordset("EH_FDATE")
    If Not IsNull(data1.Recordset("EH_TDATE")) Then dlpTo.Text = data1.Recordset("EH_TDATE")
    If Not IsNull(data1.Recordset("EH_EDATE")) Then dlpAsOf.Text = data1.Recordset("EH_EDATE")
    If Not IsNull(data1.Recordset("EH_EMP")) Then clpCode(1).Text = data1.Recordset("EH_EMP")
    If Not IsNull(data1.Recordset("EH_PT")) Then clpPT.Text = data1.Recordset("EH_PT")
    If Not IsNull(data1.Recordset("EH_HETYPE")) Then clpCode(2).Text = data1.Recordset("EH_HETYPE")
    If Not IsNull(data1.Recordset("EH_SECTION")) Then clpCode(3).Text = data1.Recordset("EH_SECTION")
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        If Not IsNull(data1.Recordset("EH_LOC")) Then clpProv.Text = data1.Recordset("EH_LOC")
    Else
        If Not IsNull(data1.Recordset("EH_LOC")) Then clpCode(4).Text = data1.Recordset("EH_LOC")
    End If
    If Not IsNull(data1.Recordset("EH_MANUAL")) Then
        chkManual.Value = data1.Recordset("EH_MANUAL")
    End If
    If Not IsNull(data1.Recordset("EH_UPDMETHOD")) Then
        txtUpdMethod.Text = data1.Recordset("EH_UPDMETHOD")
    End If
    
    Do While Not rsVE.EOF
        xOrder = rsVE("EH_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 24) Then
            If Not IsNull(rsVE("EH_BMONTH")) Then medLTServ(nOrder) = rsVE("EH_BMONTH")
            If Not IsNull(rsVE("EH_EMONTH")) Then medGTServ(nOrder) = rsVE("EH_EMONTH")
            If Not IsNull(rsVE("EH_ENTITLE")) Then medEntitle(nOrder) = rsVE("EH_ENTITLE")
            If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
                If Not IsNull(rsVE("EH_MAX")) Then medMax(nOrder) = rsVE("EH_MAX")
            End If
            If rsVE("EH_TYPE") = "D" Then optD(nOrder) = True
            If rsVE("EH_TYPE") = "H" Then optH(nOrder) = True
            If rsVE("EH_TYPE") = "F" Then optF(nOrder) = True
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close
End If

Call SET_UP_MODE

Call cmdModify_Click

End Sub

Private Sub txtUpdMethod_Change()
    If txtUpdMethod = "A" Then
        Accum.Value = True
    ElseIf txtUpdMethod = "R" Then
        Replace.Value = True
    End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
           
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT DISTINCT EH_DIV,EH_DEPT,EH_ORG,EH_EDATE,EH_FDATE,EH_TDATE,EH_EMP,EH_SECTION,EH_LOC,EH_PT,EH_HETYPE,EH_MANUAL,EH_UPDMETHOD FROM HR_HOURLYENT "
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE EH_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    If glbWFC Then 'Ticket #28553 Franks 05/03/2016
        SQLQ = SQLQ & " WHERE " & getWFCPlantSecurity("EH_SECTION")
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    data1.RecordSource = SQLQ
    data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim UpdateFlg As Boolean
    Dim Response%
    
    UpdateFlg = False

    If clpDiv.Text <> ODIV Then UpdateFlg = True
    If clpDept.Text <> ODept Then UpdateFlg = True
    If clpCode(0).Text <> oOrg Then UpdateFlg = True
    If dlpAsOf.Text <> oAsOf Then UpdateFlg = True
    If clpCode(1).Text <> oEMP Then UpdateFlg = True
    If clpPT.Text <> oEmpMode Then UpdateFlg = True
    If clpCode(2).Text <> oHETYPE Then UpdateFlg = True
    If clpCode(4).Text <> OLoc Then UpdateFlg = True
    If clpCode(3).Text <> OSection Then UpdateFlg = True
    If dlpFrom.Text <> oFDate Then UpdateFlg = True
    If dlpTo.Text <> OTDate Then UpdateFlg = True
    If chkManual.Value <> OManual Then UpdateFlg = True

    If UpdateFlg = True Then
        Response% = MsgBox("Do you want to Save changes?", MB_YESNO, "Save Changes?")    ' Get user response.
        If Response% = IDYES Then     ' Evaluate response
            'Save the changes
            Call cmdOK_Click
            Pause (0.5)
        End If
    End If

Call Display_Value
End Sub

Private Sub modSetFGlobals(strTyp$)
fglbSick% = False
fglbVac% = True
If glbCompEntVac$ = "M" Then
    fglbCompMonthly% = True
Else
    fglbCompMonthly% = False
End If
ffieldEntitle$ = "ED_VAC"
ffieldPEntitle$ = "ED_PVAC"
fglbCode$ = "VAC"

End Sub

Sub ST_OPT_VALUE()
Dim x, XoptD, XoptH, XoptF
    XoptD = optD(0).Value
    XoptH = optH(0).Value
    XoptF = optF(0).Value
    For x = 1 To 24
        optD(x).Value = XoptD
        optH(x).Value = XoptH
        optF(x).Value = XoptF
    Next
End Sub

Private Function modDelRecs(Optional xDelPrv As Boolean)
Dim BD As Integer
Dim SQLQ As String, SQL1 As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$
Dim oldEntitleUpd
Dim rsHRE As New ADODB.Recordset
Dim rzAttend As New ADODB.Recordset
Dim rsCurSal As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim pct#, prec#
Dim xKey
Dim xSkipBorrow As Boolean

modDelRecs = False

On Error GoTo modDelRecs_Err

Screen.MousePointer = HOURGLASS

If IsMissing(xDelPrv) Or xDelPrv = False Then
    Call getWSQLQ("")
    xSkipBorrow = False
ElseIf xDelPrv Then
    Call getWSQLQ("", True)
    xSkipBorrow = True
End If

SQLQ = "SELECT * FROM HRENTHRS WHERE " & fglbWSQLQ
SQLQ = SQLQ & " AND HE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
rsHRE.Open SQLQ, gdbAdoIhr001, adOpenStatic

pct# = 0
prec# = 0
If NumAddRec% = 0 Then
    prec# = 0
Else
    If rsHRE.RecordCount <> 0 Then
        prec# = 90 / rsHRE.RecordCount
    End If
End If

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0

'Ticket #22682 -Release 8.0 - Delete Previous years hourly entitlements. Assuming all the rollover is done and
'since we have Prv., the borrowing logic is accomplished there. So no need to go through this loop.
If Not xSkipBorrow Then
    Do Until rsHRE.EOF
        pct# = pct# + prec#
        MDIMain.panHelp(0).FloodPercent = pct#
        'In Vadim we will have to balance out to zero the entitlement
        'Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), Date, 0 - rsHRE("HE_ENTITLE"), "U", "Mass deleted the existing Hourly Entitlement")
        If (rsHRE("HE_PREV") + rsHRE("HE_ENTITLE")) - rsHRE("HE_TAKEN") < 0 Then
            'Used more entitlement than entitled - Jerry said to borrow it from next year
            'To borrow, add a new record in Attendance for next year
            If glbVadim Then
                'Add Record in Attendance screen
                SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & rsHRE("HE_EMPNBR")
                SQLQ = SQLQ & " AND AD_REASON = '" & rsHRE("HE_TYPE") & "'"
                SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(DateAdd("d", 1, rsHRE("HE_TDATE")))
                rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If rzAttend.EOF Then
                    rzAttend.AddNew
                End If
                rzAttend("AD_COMPNO") = "001"
                rzAttend("AD_EMPNBR") = rsHRE("HE_EMPNBR")
                rzAttend("AD_DOA") = DateAdd("d", 1, rsHRE("HE_TDATE")) 'Next year
                rzAttend("AD_REASON") = rsHRE("HE_TYPE")
                rzAttend("AD_HRS") = Abs(rsHRE("HE_PREV") + rsHRE("HE_ENTITLE") - rsHRE("HE_TAKEN")) 'Borrowed Hours
                
                SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO FROM HREMP WHERE ED_EMPNBR = " & rsHRE("HE_EMPNBR")
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsHREmp.EOF Then
                    rzAttend("AD_PAYROLL_ID") = rsHREmp("ED_PAYROLL_ID")
                    rzAttend("AD_GLNO") = rsHREmp("ED_GLNO")
                    rzAttend("AD_ORG") = rsHREmp("ED_ORG")
                End If
                rsHREmp.Close
                
                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsHRE("HE_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.BOF Then
                    If rsCurSal("SH_SALARY") > 0 Then
                        rzAttend("AD_SALARY") = rsCurSal("SH_SALARY")
                        rzAttend("AD_SALCD") = rsCurSal("SH_SALCD")
                    End If
                End If
                rsCurSal.Close
                            
                SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsHRE("HE_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.EOF Then
                    rzAttend("AD_JOB") = rsCurSal("JH_JOB")
                    rzAttend("AD_DHRS") = rsCurSal("JH_DHRS")
                    rzAttend("AD_WHRS") = rsCurSal("JH_WHRS")
                End If
                rsCurSal.Close
    
                rzAttend("AD_COMM") = "Exceeded Hours in last year so borrowed from this year"
                rzAttend("AD_LDATE") = Date
                rzAttend("AD_LUSER") = "BORROWED"
                rzAttend("AD_LTIME") = Time$
                rzAttend.Update
                rzAttend.Close
            End If
            'Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), dlpTo.Text, (rsHRE("HE_PREV") + rsHRE("HE_ENTITLE")) - rsHRE("HE_TAKEN"), "D", "Mass deleted existing Hourly Entitlement")
            Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), dlpTo.Text, 0 - (rsHRE("HE_PREV") + rsHRE("HE_ENTITLE")), "D", "Mass deleted existing Hourly Entitlement")
        Else
            'Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), dlpTo.Text, rsHRE("HE_TAKEN") - (rsHRE("HE_PREV") + rsHRE("HE_ENTITLE")), "D", "Mass deleted existing Hourly Entitlement")
            Call Append_Accrual(rsHRE("HE_EMPNBR"), rsHRE("HE_TYPE"), dlpTo.Text, 0 - (rsHRE("HE_PREV") + rsHRE("HE_ENTITLE")), "D", "Mass deleted existing Hourly Entitlement")
        End If
        
    '    'Ticket #17924 - Begin
    '    'If the Entitlement Code is suffixed with + then delete the corresponding Attendance record
    '    'for the Hourly Entitlement earned
        If Right(clpCode(2).Text, 1) = "+" Then
            'Add Record in Attendance screen
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & rsHRE("HE_EMPNBR")
            SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(2).Text & "'"
            SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFrom.Text)
            rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rzAttend.EOF Then
                rzAttend.Delete
            End If
            rzAttend.Close
            Set rzAttend = Nothing
        End If
    '    'Ticket #17924 - End
    
        xKey = rsHRE("HE_EMPNBR")
        xKey = xKey & "|" & Format(dlpFrom.Text, "dd-mmm-yyyy")
        xKey = xKey & "|" & Format(dlpTo.Text, "dd-mmm-yyyy")
        xKey = xKey & "|" & clpCode(2).Text
        xKey = xKey & "|"
        xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
        DoEvents
        Call Entitlements_Master_Integration(xKey, , True)
        DoEvents
        
        rsHRE.MoveNext
        DoEvents
    Loop
    rsHRE.Close
End If

SQLQ = "DELETE FROM HRENTHRS WHERE " & fglbWSQLQ
SQLQ = SQLQ & " AND HE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
gdbAdoIhr001.Execute SQLQ

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0

modDelRecs = True

Exit Function

modDelRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "DeleteHrEntitlement", "Delete")
modDelRecs = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsSelectionCBrant() 'Ticket #12524
'Share the logic of Sick Entitlement
Dim HEID&
Dim strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct#
Dim prec#, SQLQ As String, NumRec As Integer
Dim snapDuplic As New ADODB.Recordset
Dim oldEntitleUpd
Dim xKey
Dim xHEFromDate, xHEToDate, xDiffYear

On Error GoTo modInsSelectionCBrant_Err

modInsSelectionCBrant = False

If Not CR_SnapAddEntitle() Then Exit Function  ' create snapEntitle (form level recordset)

If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
fTablHREMP.Open "HRENTHRS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0

pct# = 0
prec# = 0
If NumAddRec% = 0 Then
    prec# = 0
Else
    prec# = 90 / NumAddRec% 'SBH avoid divid by zero...
End If

For x% = 0 To 24
    If IsNumeric(medGTServ(x%)) Then
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If Len(medLTServ(x%)) > 0 And Len(medGTServ(x%)) = 0 Then medGTServ(x%) = 9999999
Next
BeginTrans

While Not SnapAddEntitle.EOF
    pct# = pct# + prec#
    MDIMain.panHelp(0).FloodPercent = pct#

    
    If IsNull(SnapAddEntitle(fglbWDateS$)) Then
        GoTo lblNext2Rec
    End If

    varStartDate = SnapAddEntitle(fglbWDateS$)   ' set start date
    xDiffYear = DateDiff("d", varStartDate, Now) / 365
    
    If xDiffYear > 1 Then
        xHEFromDate = DateAdd("YYYY", CInt(xDiffYear), varStartDate)
    Else
        xHEFromDate = varStartDate
    End If
    
    If xHEFromDate > Now Then
        xHEFromDate = DateAdd("YYYY", -1, xHEFromDate)
    End If
    xHEToDate = DateAdd("YYYY", 1, xHEFromDate)
    xHEToDate = DateAdd("d", -1, xHEToDate)
    
    If Not IsNumeric(SnapAddEntitle("JH_DHRS")) Then
        dblDHours# = 0
    Else
        dblDHours# = SnapAddEntitle("JH_DHRS")
    End If
    If Not IsNumeric(SnapAddEntitle("JH_FTENUM")) Then
        dblFTEHours# = 0
    Else
        dblFTEHours# = SnapAddEntitle("JH_FTENUM")
    End If
    
    'dblServiceYears# = (DateDiff("d", varStartDate, Now) / 365) * 12
    'dblServiceYears# = MonthDiff(CVDate(varStartDate), Date)
    dblServiceYears# = MonthDiff(CVDate(varStartDate), dlpAsOf.Text)    'Ticket #17924
    
    If dblServiceYears# < 0 Then GoTo lblNext2Rec     'laura 03/06/98
    
    intWhereFit& = -1   ' first record can be just less than
    For x% = 0 To 24
        If medLTServ(x%) = "" And medGTServ(x%) = "" Then Exit For
        If IsNumeric(Val(medLTServ(x%))) And medGTServ(x%) = "" Then
            If dblServiceYears# >= CDbl(Val(medLTServ(x%))) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
        If IsNumeric(medLTServ(x%)) And IsNumeric(medGTServ(x%)) Then
            If dblServiceYears# >= CDbl(Val(medLTServ(x%))) And dblServiceYears# <= CDbl(Val(medGTServ(x%))) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNext2Rec  ' skip record if not in any of the ranges
    dblNewEntitle# = Val(medEntitle(intWhereFit&))   'laura
    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
        dblNewEntitle# = dblNewEntitle# * dblDHours#
    End If
    If optH(intWhereFit&) = True Then           ' Entitlements entered in Hours
        dblNewEntitle# = dblNewEntitle#
    End If
    If optF(intWhereFit&) = True Then           ' Entitlements entered in FTE
        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
    End If

    SQLQ = "SELECT HE_EMPNBR,HE_TYPE,HE_ID ,"
    SQLQ = SQLQ & " HE_ENTITLE, HE_TDATE FROM HRENTHRS "
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(xHEToDate)  'dlpTo.Text
    snapDuplic.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not snapDuplic.EOF And Not snapDuplic.BOF Then
'        xID = snapDuplic("HE_ID")
        snapDuplic.MoveLast
    End If

    NumRec = snapDuplic.RecordCount
    If snapDuplic.EOF Then
        oldEntitleUpd = 0
    Else
        oldEntitleUpd = snapDuplic("HE_ENTITLE")
    End If
    If Accum = True Then
        If NumRec > 0 Then
            dblEntitleUpd = snapDuplic("HE_ENTITLE")
        Else
            dblEntitleUpd = 0
        End If
    Else
        dblEntitleUpd = 0
    End If
    snapDuplic.Close
    
    If Accum = True Then
        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
    Else
        dblEntitleUpd = dblNewEntitle
    End If

    DtTm = Now
    
If Accum = True Then
    If NumRec > 0 Then  'if accumulate and found duplicate record
    
        SQLQ = "UPDATE HRENTHRS "
        SQLQ = SQLQ & " SET HE_ENTITLE = " & dblEntitleUpd & " "
        SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
        SQLQ = SQLQ & " AND HRENTHRS.HE_TYPE = '" & clpCode(2).Text & "' "
        SQLQ = SQLQ & " AND HRENTHRS.HE_TDATE = " & Date_SQL(xHEToDate)
        
        gdbAdoIhr001.Execute (SQLQ)
        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, Date, dblEntitleUpd - oldEntitleUpd, "U", "Mass changed the existing Hourly Entitlement")
    Else
        fTablHREMP.AddNew     'if accumulate and no duplicate record
        fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
        fTablHREMP("HE_COMPNO") = "001"
        fTablHREMP("HE_TYPE_TABL") = "ADRE"
        fTablHREMP("HE_TYPE") = clpCode(2).Text
        fTablHREMP("HE_FDATE") = xHEFromDate 'dlpFrom.Text
        fTablHREMP("HE_TDATE") = xHEToDate  'dlpTo.Text
        fTablHREMP("HE_ENTITLE") = dblEntitleUpd
        fTablHREMP("HE_COE") = True
        fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
        fTablHREMP("HE_LDATE") = Now
        fTablHREMP("HE_LTIME") = Time$
        fTablHREMP("HE_LUSER") = glbUserID
        fTablHREMP.Update
        '    xID = fTablHREMP("HE_ID")
        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, xHEFromDate, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")

    End If
Else
    SQLQ$ = "DELETE FROM HRENTHRS "
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(xHEToDate)
    
    gdbAdoIhr001.Execute SQLQ
    
    fTablHREMP.AddNew
    
    fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
    fTablHREMP("HE_COMPNO") = "001"
    fTablHREMP("HE_TYPE_TABL") = "ADRE"
    fTablHREMP("HE_TYPE") = clpCode(2).Text
    fTablHREMP("HE_FDATE") = xHEFromDate
    fTablHREMP("HE_TDATE") = xHEToDate
    fTablHREMP("HE_ENTITLE") = dblEntitleUpd
    fTablHREMP("HE_COE") = True
    fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
    fTablHREMP("HE_LDATE") = Now
    fTablHREMP("HE_LTIME") = Time$
    fTablHREMP("HE_LUSER") = glbUserID
    fTablHREMP.Update
    '    xID = fTablHREMP("HE_ID")
    If NumRec > 0 Then  'if accumulate and found duplicate record
        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, xHEFromDate, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
    Else
        Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, xHEFromDate, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
    End If

End If
    DoEvents
    xKey = SnapAddEntitle("ED_EMPNBR")
    xKey = xKey & "|" & Format(xHEFromDate, "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(xHEToDate, "dd-mmm-yyyy")
    xKey = xKey & "|" & clpCode(2).Text
    xKey = xKey & "|" & dblEntitleUpd
    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, 0)
    DoEvents
lblNext2Rec:
    SnapAddEntitle.MoveNext
Wend

modInsSelectionCBrant = True

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0

CommitTrans

fTablHREMP.Close

SnapAddEntitle.Close

Screen.MousePointer = DEFAULT

Exit Function

modInsSelectionCBrant_Err:

If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
   'MsgBox "Conflicting Dates"
    Screen.MousePointer = DEFAULT
    Exit Function
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InsertEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsSelection()
'laura 03/04/98
Dim dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct#
Dim prec#, SQLQ As String, NumRec As Integer
Dim snapDuplic As New ADODB.Recordset
Dim rzAttend As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim rsCurJobSal As New ADODB.Recordset
Dim oldEntitleUpd
Dim xKey
Dim rsJOB As New ADODB.Recordset
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012
Dim dblNewMax# 'Ticket #27729 Franks 03/14/2016 Carizon

On Error GoTo modInsSelection_Err

modInsSelection = False

If Not CR_SnapAddEntitle() Then Exit Function  ' create snapEntitle (form level recordset)


'Ticket #22682 - Release 8.0: Check Accrual File to see if the update already done for Monthly Updates only. This is
'to avoid multiple updates for the same month.
'Only for Monthly updates
'If glbCompEntSick$ = "M" Then
Do While Not SnapAddEntitle.EOF
    If Accrual_Rec_Exists(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, "'U','A'") Then
        Response% = MsgBox("'Update Entitlement' already done for at least 1 employee in this selection for the Effective Date: " & dlpAsOf.Text & "." & Chr(10) & Chr(10) & "Are you sure you want to proceed with this Update?", vbExclamation + vbYesNo, "Update Entitlements")
        If Response% = IDNO Then
            Exit Function
        End If
        
        Exit Do
    End If
    
    SnapAddEntitle.MoveNext
    DoEvents
Loop
'End If

SnapAddEntitle.MoveFirst
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5


If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
fTablHREMP.Open "HRENTHRS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
pct# = 0
prec# = 0
If NumAddRec% = 0 Then
    prec# = 0
Else
    prec# = 90 / NumAddRec% 'SBH avoid divid by zero...
End If
For x% = 0 To 24
    If IsNumeric(medGTServ(x%)) Then
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If Len(medLTServ(x%)) > 0 And Len(medGTServ(x%)) = 0 Then medGTServ(x%) = 9999999
Next
BeginTrans

While Not SnapAddEntitle.EOF
    pct# = pct# + prec#
    MDIMain.panHelp(0).FloodPercent = pct#

    
    If IsNull(SnapAddEntitle(fglbWDate$)) Then
        GoTo lblNext2Rec
    End If
    
    'clpCode(2).Text
    'Ticket #18518 "+" and "-" need the Hourly entitlement setup before do the attendance import
    'but it cannot update "VAC" and "SICK"
    'Lanark
    'If glbCompSerial = "S/N - 2172W" Then
    'Ticket #19782 Franks 02/03/2011 for Frontenac
    If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
        If clpCode(2).Text = "VAC" Or clpCode(2).Text = "SICK" Then
             GoTo lblNext2Rec
        End If
    End If

    varStartDate = SnapAddEntitle(fglbWDate$)  ' set start date
    If Not IsNumeric(SnapAddEntitle("JH_DHRS")) Then
        dblDHours# = 0
    Else
        dblDHours# = SnapAddEntitle("JH_DHRS")
    End If
    If Not IsNumeric(SnapAddEntitle("JH_FTENUM")) Then
        dblFTEHours# = 0
    Else
        dblFTEHours# = SnapAddEntitle("JH_FTENUM")
    End If
    
    'dblServiceYears# = (DateDiff("d", varStartDate, Now) / 365) * 12
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(dlpAsOf.Text))    'Ticket #17924
    If dblServiceYears# < 0 Then GoTo lblNext2Rec     'laura 03/06/98
    intWhereFit& = -1   ' first record can be just less than
    For x% = 0 To 24
        If medLTServ(x%) = "" And medGTServ(x%) = "" Then Exit For
        If IsNumeric(medLTServ(x%)) And medGTServ(x%) = "" Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
        If IsNumeric(medLTServ(x%)) And IsNumeric(medGTServ(x%)) Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNext2Rec  ' skip record if not in any of the ranges
    
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & SnapAddEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly

    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
        xTotEmpHours = 0
        Do While Not rsJOB.EOF
            If optD(intWhereFit&) = True Then  ' Entitlements entered in days
                If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
            End If
            If optF(intWhereFit&) = True Then  ' FTE
                If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                    xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                End If
            End If
            rsJOB.MoveNext
        Loop
    End If
    rsJOB.Close
    
    dblNewMax# = 0
    
    If glbFrench Then
        dblNewEntitle# = CDbl(medEntitle(intWhereFit&))   'laura
    Else
        dblNewEntitle# = Val(medEntitle(intWhereFit&))   'laura
    End If
    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
        'Ticket #22766 - KidsLink - sum up the FTE for multi positions
        If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
            dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * xTotEmpHours
            End If
        Else
            dblNewEntitle# = dblNewEntitle# * dblDHours#
            If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
            End If
        End If
    End If
    If optH(intWhereFit&) = True Then           ' Entitlements entered in Hours
        dblNewEntitle# = dblNewEntitle#
        If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
            If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&)
        End If
    End If
    If optF(intWhereFit&) = True Then           ' Entitlements entered in FTE
        'Ticket #22766 - KidsLink - sum up the FTE for multi positions
        If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
            dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * xTotEmpHours
            End If
        Else
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
                If medMax(intWhereFit&) <> 0 Then dblNewMax# = medMax(intWhereFit&) * dblDHours#
            End If
        End If
    End If
    If medMax(0).Visible = True Then 'Ticket #27729 Franks 03/14/2016 Carizon
        If dblNewMax# > 0 Then
            If dblNewEntitle# > dblNewMax# Then
                dblNewEntitle# = dblNewMax#
            End If
        End If
    End If
    
    SQLQ = "SELECT HE_EMPNBR,HE_TYPE,HE_ID ,"
    SQLQ = SQLQ & " HE_ENTITLE, HE_TDATE FROM HRENTHRS "
    SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
    SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
    SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
    snapDuplic.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    
    If Not snapDuplic.EOF And Not snapDuplic.BOF Then
'        xID = snapDuplic("HE_ID")
        snapDuplic.MoveLast
    End If

    NumRec = snapDuplic.RecordCount
    If snapDuplic.EOF Then
        oldEntitleUpd = 0
    Else
        oldEntitleUpd = snapDuplic("HE_ENTITLE")
    End If
    If Accum = True Then
        If NumRec > 0 Then
            dblEntitleUpd = snapDuplic("HE_ENTITLE")
        Else
            dblEntitleUpd = 0
        End If
    Else
        dblEntitleUpd = 0
    End If
    snapDuplic.Close
    
    If Accum = True Then
        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
    Else
        dblEntitleUpd = dblNewEntitle
    End If

    DtTm = Now
    
    If Accum = True Then
        If NumRec > 0 Then  'if accumulate and found duplicate record
            SQLQ = "UPDATE HRENTHRS "
            SQLQ = SQLQ & " SET HE_ENTITLE = " & dblEntitleUpd & " "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
            SQLQ = SQLQ & " AND HRENTHRS.HE_TYPE = '" & clpCode(2).Text & "' "
            SQLQ = SQLQ & " AND HRENTHRS.HE_TDATE = " & Date_SQL(dlpTo.Text)
            
            gdbAdoIhr001.Execute (SQLQ)
            Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass changed the existing Hourly Entitlement") 'Ticket #17924
        Else
            'Ticket #17924 - If Flex logic (+) then update the existing Flex code hourly entitlement record instead
            'of adding a new record.
            If Right(clpCode(2).Text, 1) = "+" Then
                If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
                SQLQ = "SELECT * FROM HRENTHRS "
                SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
                SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
                SQLQ = SQLQ & " ORDER BY HE_FDATE DESC"
                fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not fTablHREMP.EOF Then
                    fTablHREMP.MoveFirst
                Else
                    fTablHREMP.AddNew
                    fTablHREMP("HE_PREV") = 0
                End If
            Else
                fTablHREMP.AddNew     'if accumulate and no duplicate record
                fTablHREMP("HE_PREV") = 0
            End If
            
            fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
            fTablHREMP("HE_COMPNO") = "001"
            fTablHREMP("HE_TYPE_TABL") = "ADRE"
            fTablHREMP("HE_TYPE") = clpCode(2).Text
            fTablHREMP("HE_FDATE") = dlpFrom.Text
            fTablHREMP("HE_TDATE") = dlpTo.Text
            fTablHREMP("HE_ENTITLE") = dblEntitleUpd
            fTablHREMP("HE_COE") = True
            fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
            fTablHREMP("HE_LDATE") = Now
            fTablHREMP("HE_LTIME") = Time$
            fTablHREMP("HE_LUSER") = glbUserID
            fTablHREMP.Update
            '    xID = fTablHREMP("HE_ID")
            'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
            Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
        End If
    Else
                
        'Ticket #17924 - If Flex logic (+) then update the existing Flex code hourly entitlement record instead
        'of adding a new record.
        If Right(clpCode(2).Text, 1) = "+" Then
            If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
            SQLQ = "SELECT * FROM HRENTHRS "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
            SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
            SQLQ = SQLQ & " ORDER BY HE_FDATE DESC"
            fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not fTablHREMP.EOF Then
                fTablHREMP.MoveFirst
            Else
                fTablHREMP.AddNew
                fTablHREMP("HE_PREV") = 0
            End If
        Else
            'Ticket #18559 - Jerry does not want the Previous to be replaced with 0 after the rollover which
            'creates a new record on the Hourly Entitlement screen. In which case we cannot delete an existing
            'Hourly Entitlement record but instead update the values.
            'SQLQ$ = "DELETE FROM HRENTHRS "
            'SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
            'SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
            'SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
            'gdbAdoIhr001.Execute SQLQ
            
            If fTablHREMP.State <> adStateClosed Then fTablHREMP.Close
            SQLQ = "SELECT * FROM HRENTHRS "
            SQLQ = SQLQ & " WHERE HE_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
            SQLQ = SQLQ & " AND HE_TYPE = '" & clpCode(2).Text & "'"
            SQLQ = SQLQ & " AND HE_TDATE = " & Date_SQL(dlpTo.Text)
            fTablHREMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If fTablHREMP.EOF Then
                fTablHREMP.AddNew
                fTablHREMP("HE_PREV") = 0
            End If
        End If
        
        'fTablHREMP.AddNew
        
        fTablHREMP("HE_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
        fTablHREMP("HE_COMPNO") = "001"
        fTablHREMP("HE_TYPE_TABL") = "ADRE"
        fTablHREMP("HE_TYPE") = clpCode(2).Text
        fTablHREMP("HE_FDATE") = dlpFrom.Text
        fTablHREMP("HE_TDATE") = dlpTo.Text
        fTablHREMP("HE_ENTITLE") = dblEntitleUpd
        fTablHREMP("HE_COE") = True
        fTablHREMP("HE_DHRS") = SnapAddEntitle("ED_DHRS")
        fTablHREMP("HE_LDATE") = Now
        fTablHREMP("HE_LTIME") = Time$
        fTablHREMP("HE_LUSER") = glbUserID
        fTablHREMP.Update
        '    xID = fTablHREMP("HE_ID")
        If NumRec > 0 Then  'if accumulate and found duplicate record
            'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
            Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
        Else
            'Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
            Call Append_Accrual(SnapAddEntitle("ED_EMPNBR"), clpCode(2).Text, dlpAsOf.Text, dblEntitleUpd, "A", "Mass added the Hourly Entitlement")
        End If
    End If

    'Ticket #22682 - Release 8.0: Jerry said not to check for duplicate, simply add new Attendance record, even
    'though it is a duplicate record.
'    'Ticket #17924 - Begin
'    'If the Entitlement Code is suffixed with + then insert an Attendance record
'    'for the Hourly Entitlement earned - helps in the Recalculate function
    If Right(clpCode(2).Text, 1) = "+" Then
        'Add Record in Attendance screen
        'Ticket #22682 - Release 8.0: Do not check for duplicates
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
        'SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & SnapAddEntitle("ED_EMPNBR")
        'SQLQ = SQLQ & " AND AD_REASON = '" & clpCode(2).Text & "'"
        'Ticket #18550 - Attendance record date cannot be prior to hire date
        'If CVDate(SnapAddEntitle("ED_DOH")) > CVDate(dlpFrom.Text) Then
        '    SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(SnapAddEntitle("ED_DOH"))
        'Else
        '    SQLQ = SQLQ & " AND AD_DOA =" & Date_SQL(dlpFrom.Text)
        'End If
        rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'If rzAttend.EOF Then
            rzAttend.AddNew
        'End If
        rzAttend("AD_COMPNO") = "001"
        rzAttend("AD_EMPNBR") = SnapAddEntitle("ED_EMPNBR")
        rzAttend("AD_DOA") = dlpFrom.Text
        rzAttend("AD_REASON") = clpCode(2).Text
        rzAttend("AD_HRS") = dblEntitleUpd

        SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO,ED_DOH FROM HREMP WHERE ED_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
        rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsHREmp.EOF Then
            rzAttend("AD_PAYROLL_ID") = rsHREmp("ED_PAYROLL_ID")
            rzAttend("AD_GLNO") = rsHREmp("ED_GLNO")
            rzAttend("AD_ORG") = rsHREmp("ED_ORG")
            
            'Ticket #18550 - Attendance record date cannot be prior to hire date
            If CVDate(rsHREmp("ED_DOH")) > CVDate(dlpFrom.Text) Then
                rzAttend("AD_DOA") = rsHREmp("ED_DOH")
            End If
        End If
        rsHREmp.Close

        SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
        rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCurJobSal.BOF Then
            If rsCurJobSal("SH_SALARY") > 0 Then
                rzAttend("AD_SALARY") = rsCurJobSal("SH_SALARY")
                rzAttend("AD_SALCD") = rsCurJobSal("SH_SALCD")
            End If
        End If
        rsCurJobSal.Close
        Set rsCurJobSal = Nothing

        SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & SnapAddEntitle("ED_EMPNBR")
        rsCurJobSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCurJobSal.EOF Then
            rzAttend("AD_JOB") = rsCurJobSal("JH_JOB")
            rzAttend("AD_DHRS") = rsCurJobSal("JH_DHRS")
            rzAttend("AD_WHRS") = rsCurJobSal("JH_WHRS")
        End If
        rsCurJobSal.Close
        Set rsCurJobSal = Nothing

        'Ticket #18550
        'rzAttend("AD_COMM") = "Entitlement earned for the period: " & dlpFrom.Text & " to " & dlpTo.Text & "."
        rzAttend("AD_COMM") = "Entitlement earned for the period: " & rzAttend("AD_DOA") & " to " & dlpTo.Text & "."
        rzAttend("AD_LDATE") = Date
        rzAttend("AD_LUSER") = glbUserID
        rzAttend("AD_LTIME") = Time$
        rzAttend.Update
        rzAttend.Close
    End If
'    'Ticket #17924 - End
    
    DoEvents
    xKey = SnapAddEntitle("ED_EMPNBR")
    xKey = xKey & "|" & Format(dlpFrom.Text, "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpTo.Text, "dd-mmm-yyyy")
    xKey = xKey & "|" & clpCode(2).Text
    xKey = xKey & "|" & dblEntitleUpd
    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, 0)
    DoEvents
    
lblNext2Rec:
    SnapAddEntitle.MoveNext
Wend

modInsSelection = True

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0

CommitTrans

fTablHREMP.Close

SnapAddEntitle.Close

Screen.MousePointer = DEFAULT

Exit Function

modInsSelection_Err:

If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
   'MsgBox "Conflicting Dates"
    Screen.MousePointer = DEFAULT
    Exit Function
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InsertEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If


End Function

Private Function modUpdateSelection()
Dim HEID&
Dim strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec%
Dim rsHE As New ADODB.Recordset
Dim oldEntitleUpd
Dim xKey
Dim rsJOB As New ADODB.Recordset
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012

On Error GoTo modUpdateSelection_Err

modUpdateSelection = False

If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)

Screen.MousePointer = DEFAULT
If snapEntitle.BOF And snapEntitle.EOF Then
    MsgBox "Employees for this selection do not exist!"
    Exit Function
Else
    lngRecs& = snapEntitle.RecordCount
'    Msg$ = lngRecs& & " Records to process" & Chr(10) & "Would You Like To Proceed?"
'    Title$ = "Update Entitlements"
'    DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
'    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
'    If Response% = IDNO Then    ' Evaluate response
'        Exit Function
End If
Screen.MousePointer = HOURGLASS
'End If
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 5
For x% = 0 To 24
    If IsNumeric(medGTServ(x%)) Then
        If glbFrench Then
            If medGTServ(x%) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        Else
            If Val(medGTServ(x%)) = Int(medGTServ(x%)) Then medGTServ(x%) = medGTServ(x%) + 0.99
        End If
    End If
    If Len(medLTServ(x%)) > 0 And Len(medGTServ(x%)) = 0 Then medGTServ(x%) = 9999999
Next
BeginTrans

While Not snapEntitle.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / lngRecs&))
    MDIMain.panHelp(0).FloodPercent = pct%

    HEID& = snapEntitle("HE_ID")
    oldEntitleUpd = snapEntitle("HE_ENTITLE")
    
    If Accum = True Then
        dblEntitleUpd = snapEntitle("HE_ENTITLE")
    Else
        dblEntitleUpd = 0
    End If
    
    spt = snapEntitle("ED_PT")
    strDivision$ = snapEntitle("ED_DIV")

    If IsNull(snapEntitle(fglbWDate$)) Then
        GoTo lblNextRec
    End If

    varStartDate = snapEntitle(fglbWDate$)
    
    If Not IsNumeric(snapEntitle("JH_DHRS")) Then
        dblDHours# = 0
    Else
        dblDHours# = snapEntitle("JH_DHRS")
    End If
    If Not IsNumeric(snapEntitle("JH_FTENUM")) Then
        dblFTEHours# = 0
    Else
        dblFTEHours# = snapEntitle("JH_FTENUM")
    End If
    
    'dblServiceYears# = (DateDiff("d", varStartDate, Now) / 365) * 12
    dblServiceYears# = MonthDiff(CVDate(varStartDate), Date)
    
    intWhereFit& = -1   ' first record can be just less than
    For x% = 0 To 24
        If medLTServ(x%) = "" And Not medGTServ(x%) = "" Then Exit Function

        If IsNumeric(medLTServ(x%)) And medGTServ(x%) = "" Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) Then
                intWhereFit& = x%
                Exit For
            End If
        End If

        If IsNumeric(medLTServ(x%)) And IsNumeric(medGTServ(x%)) Then
            If dblServiceYears# >= CDbl(medLTServ(x%)) And dblServiceYears# <= CDbl(medGTServ(x%)) Then
                intWhereFit& = x%
                Exit For
            End If
        End If
    Next x%
    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges

    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly

    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
        xTotEmpHours = 0
        Do While Not rsJOB.EOF
            If optD(intWhereFit&) = True Then  ' Entitlements entered in days
                If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
            End If
            If optF(intWhereFit&) = True Then  ' FTE
                If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                    xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                End If
            End If
            rsJOB.MoveNext
        Loop
    End If
    rsJOB.Close

    dblNewEntitle# = medEntitle(intWhereFit&)
    If optD(intWhereFit&) = True Then           ' Entitlements entered in days
        'Ticket #22766 - KidsLink - sum up the FTE for multi positions
        If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then   'Kerrys Place Ticket #21843 Franks 04/12/2012
            dblNewEntitle# = dblNewEntitle# * xTotEmpHours
        Else
            dblNewEntitle# = dblNewEntitle# * dblDHours#
        End If
    End If
    If optH(intWhereFit&) = True Then           ' Entitlements entered in Hours
        dblNewEntitle# = dblNewEntitle#
    End If
    If optF(intWhereFit&) = True Then           ' Entitlements entered in FTE
        'Ticket #22766 - KidsLink - sum up the FTE for multi positions
        ' (Entitlement * Hrs/Day) * FTE Factor
        If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
            dblNewEntitle# = dblNewEntitle# * xTotEmpHours
        Else
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
        End If
    End If
    If Accum = True Then
        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
    Else
        dblEntitleUpd = dblNewEntitle
    End If
    
    DtTm = Now
        
    rsHE.Open "SELECT * FROM HRENTHRS WHERE HE_ID= " & HEID&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsHE("HE_ENTITLE") = dblEntitleUpd
    rsHE("HE_LDATE") = Now
    rsHE("HE_LTIME") = Time$
    rsHE("HE_LUSER") = glbUserID
    rsHE.Update
    rsHE.Close
    
    If Accum = True Then
        Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, Date, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
    Else
        Call Append_Accrual(snapEntitle("ED_EMPNBR"), clpCode(2).Text, dlpFrom.Text, dblEntitleUpd - oldEntitleUpd, "U", "Mass modified the Hourly Entitlement")
    End If
    
    DoEvents
    xKey = snapEntitle("HE_EMPNBR")
    xKey = xKey & "|" & Format(dlpFrom.Text, "dd-mmm-yyyy")
    xKey = xKey & "|" & Format(dlpTo.Text, "dd-mmm-yyyy")
    xKey = xKey & "|" & clpCode(2).Text
    xKey = xKey & "|" & dblEntitleUpd
    xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
    Call Entitlements_Master_Integration(xKey, HEID&)
    
    DoEvents
lblNextRec:
    snapEntitle.MoveNext

Wend
modUpdateSelection = True
MDIMain.panHelp(0).FloodType = 0
CommitTrans

'fTablHREMP.Close

snapEntitle.Close

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:

If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
   'MsgBox "Conflicting Dates"
    Screen.MousePointer = DEFAULT
    Exit Function
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Function getWSQLQ(xType, Optional xDelPrv As Boolean)
Dim SQLQ As String
Dim xDiv, xDept, xORG, xFDate, xTDate, xEMP, xEmpMode, xHETYPE
Dim xLoc, xSection

fglbESQLQ = glbSeleDeptUn

If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & clpCode(1).Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(clpProv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PROV = '" & clpProv.Text & "' "
Else
    If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(4).Text & "' "
End If
If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & clpPT.Text & "' "

fglbWSQLQ = "HE_TYPE = '" & clpCode(2).Text & "' "

'Ticket #18518 "+" and "-" need the Hourly entitlement setup before do the attendance import
'but it cannot update "VAC" and "SICK"
'Lanark Ticket #17711
'If glbCompSerial = "S/N - 2172W" Then
'Ticket #19782 Franks 02/03/2011 for Frontenac
If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
    fglbWSQLQ = fglbWSQLQ & " AND NOT (HE_TYPE = 'VAC' OR HE_TYPE = 'SICK') "
End If

'Ticket #22682 - Release 8.0 - Option to Delete Previous Year's Hourly Entitlements
If Not glbCBrant And (IsMissing(xDelPrv) Or xDelPrv = False) Then
    fglbWSQLQ = fglbWSQLQ & " AND HE_FDATE >= " & Date_SQL(dlpFrom.Text)
    fglbWSQLQ = fglbWSQLQ & " AND HE_TDATE <= " & Date_SQL(dlpTo.Text)
ElseIf Not glbCBrant And Not IsMissing(xDelPrv) Then
    If xDelPrv Then
        'Delete all Hourly Entitlements records that are prior to From Date of currently selected Hourly
        'Entitlement record
        fglbWSQLQ = fglbWSQLQ & " AND HE_TDATE < " & Date_SQL(dlpFrom.Text)
    End If
End If

If xType = "" Then Exit Function

If xType = "O" Then
    xDiv = ODIV
    xDept = ODept
    xORG = oOrg
    xFDate = oFDate
    xTDate = OTDate
    xEMP = oEMP
    xEmpMode = oEmpMode
    xHETYPE = oHETYPE
    xLoc = OLoc
    xSection = OSection
ElseIf xType = "D" Then
    xDiv = data1.Recordset("EH_DIV")
    xDept = data1.Recordset("EH_DEPT")
    xORG = data1.Recordset("EH_ORG")
    xFDate = data1.Recordset("EH_FDATE")
    xTDate = data1.Recordset("EH_TDATE")
    xEMP = data1.Recordset("EH_EMP")
    xEmpMode = data1.Recordset("EH_PT")
    xHETYPE = data1.Recordset("EH_HETYPE")
    xLoc = data1.Recordset("EH_LOC")
    xSection = data1.Recordset("EH_SECTION")
Else
    xDiv = clpDiv.Text
    xDept = clpDept.Text
    xORG = clpCode(0).Text
    xFDate = dlpFrom.Text
    xTDate = dlpTo.Text
    xEMP = clpCode(1).Text
    xEmpMode = clpPT.Text
    xHETYPE = clpCode(2).Text
    If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
        xLoc = clpProv.Text
    Else
        xLoc = clpCode(4).Text
    End If
    xSection = clpCode(3).Text
End If
    
If Len(xDiv) = 0 Or IsNull(xDiv) Then
    fglbVSQLQ = " (EH_DIV IS NULL OR EH_DIV='')"
Else
    fglbVSQLQ = "EH_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Or IsNull(xDept) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_DEPT IS NULL OR EH_DEPT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_DEPT = '" & xDept & "'"
End If
If Len(xORG) = 0 Or IsNull(xORG) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_ORG IS NULL OR EH_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_ORG = '" & xORG & "'"
End If
If Len(oFDate) > 0 Or IsNull(xFDate) Then
    SQLQ = SQLQ & " AND  EH_FDATE = " & Date_SQL(oFDate)
End If
If Len(OTDate) > 0 Or IsNull(xTDate) Then
    SQLQ = SQLQ & " AND  EH_TDATE = " & Date_SQL(OTDate)
End If
If Len(xEMP) = 0 Or IsNull(xEMP) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_EMP IS NULL OR EH_EMP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_EMP = '" & xEMP & "'"
End If
If Len(xLoc) = 0 Or IsNull(xLoc) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_LOC IS NULL OR EH_LOC='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_LOC = '" & xLoc & "'"
End If
If Len(xSection) = 0 Or IsNull(xSection) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_SECTION IS NULL OR EH_SECTION='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_SECTION = '" & xSection & "'"
End If
If Len(xEmpMode) = 0 Or IsNull(xEmpMode) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_PT IS NULL OR EH_PT='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_PT = '" & xEmpMode & "' "
End If
If Len(xHETYPE) = 0 Or IsNull(xHETYPE) Then
    fglbVSQLQ = fglbVSQLQ & " AND (EH_HETYPE IS NULL OR EH_HETYPE='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND EH_HETYPE = '" & xHETYPE & "'"
End If

End Function

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
    cmdPrintAll.Enabled = False
    cmdUpdate.Enabled = False
    cmdUpdateAll.Enabled = False
    'cmdAddEnt.Enabled = False
    cmdDeleteEnt.Enabled = False
ElseIf Me.data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = False
    cmdUpdate.Enabled = False
    cmdUpdateAll.Enabled = False
    'cmdAddEnt.Enabled = False
    cmdDeleteEnt.Enabled = False
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
    cmdUpdate.Enabled = True
    cmdUpdateAll.Enabled = True
    'cmdAddEnt.Enabled = True
    cmdDeleteEnt.Enabled = True
End If

'Ticket #18518 "+" and "-" need the Hourly entitlement setup before do the attendance import
'but it cannot update "VAC" and "SICK"
'Lanark Ticket #17711
'They keep Entitlements in GP, we import the Ent and taken,
'info:HR can not do Ent update, just use Rule to get date range
'Ticket #19782 Franks 02/03/2011 for Frontenac
If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
    cmdUpdate.Enabled = False
    cmdUpdateAll.Enabled = False
End If

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

Call ST_UPD_MODE(TF)

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
UpdateRight = gSec_Upd_Hrly_Entitlements
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

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    Call getWSQLQ("")
    
    SQLQ = "SELECT COUNT(DISTINCT ED_EMPNBR) AS TOT_REC "
    If glbOracle Then
        SQLQ = SQLQ & "FROM HREMP, HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT<>0"
    Else
        SQLQ = SQLQ & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
    End If
    
    SQLQ = SQLQ & " AND " & fglbESQLQ
    
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        recCount = rsEmp("TOT_REC")
    Else
        recCount = 0
    End If
    rsEmp.Close
    Set rsEmp = Nothing
    
    getRecordCount_Add = recCount

End Function

Private Function getRecordCount_Modify()
    Dim SQLQ As String
    Dim rsHRE As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Modify = 0
    recCount = 0

    Call getWSQLQ("")
    
    SQLQ = "SELECT COUNT(HE_EMPNBR) AS TOT_REC FROM qry_MU_Hourly "
    SQLQ = SQLQ & " WHERE " & fglbESQLQ & " AND " & fglbWSQLQ
    rsHRE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHRE.EOF Then
        recCount = rsHRE("TOT_REC")
    Else
        recCount = 0
    End If
    rsHRE.Close
    Set rsHRE = Nothing
    
    getRecordCount_Modify = recCount

End Function

Private Function getRecordCount_Delete(xDelPrv As Boolean)
    Dim SQLQ As String
    Dim rsHRE As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Delete = 0
    recCount = 0

    If xDelPrv Then
        Call getWSQLQ("", True)
    Else
        Call getWSQLQ("")
    End If
    
    SQLQ = "SELECT COUNT(HE_EMPNBR) AS TOT_REC FROM HRENTHRS WHERE " & fglbWSQLQ
    SQLQ = SQLQ & " AND HE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
    rsHRE.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHRE.EOF Then
        recCount = rsHRE("TOT_REC")
    Else
        recCount = 0
    End If
    rsHRE.Close
    Set rsHRE = Nothing
    
    getRecordCount_Delete = recCount

End Function

Private Sub ScreenSetup(xVis As Boolean)
Dim I As Integer
    lblHeading(2).Top = 4620
    lblHeading(2).Visible = xVis
    lblHeading(3).Visible = False
    For I = 0 To 24
        medMax(I).Visible = xVis
    Next
    
    'Ticket #29617 - Mississaugas of Scugog Island First Nation
    If glbCompSerial = "S/N - 2485W" Then
        lblHeading(2).Top = 4380
        lblHeading(3).Top = 4620    'Pay Period
        lblHeading(2).Caption = "Maximum /"
        lblHeading(3).Visible = xVis
    
        optD(0).Enabled = False
        'optH(0).Enabled = False
        optF(0).Enabled = False
    End If
End Sub
