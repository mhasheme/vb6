VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEHSINJURYWF7 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Injury/Location for WSIB Form 7"
   ClientHeight    =   10365
   ClientLeft      =   330
   ClientTop       =   810
   ClientWidth     =   12750
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10365
   ScaleWidth      =   12750
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   6735
      LargeChange     =   315
      Left            =   12360
      Max             =   100
      SmallChange     =   315
      TabIndex        =   196
      Top             =   2880
      Width           =   340
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   8760
      Top             =   9600
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Caption         =   "Ado2"
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsinjrWF7.frx":0000
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "fehsinjrWF7.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   11895
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12750
      _Version        =   65536
      _ExtentX        =   22490
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
         Left            =   6960
         TabIndex        =   11
         Top             =   135
         Width           =   1305
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
         Left            =   120
         TabIndex        =   7
         Top             =   160
         Width           =   1005
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   6
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   5
         Top             =   135
         Width           =   720
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2760
      MaxLength       =   25
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   10110
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4560
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   10110
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EC_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10110
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   9705
      Width           =   12750
      _Version        =   65536
      _ExtentX        =   22490
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
      Begin VB.CommandButton cmdF7Sections 
         Appearance      =   0  'Flat
         Caption         =   "Additional Form 7 Sections"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4485
         TabIndex        =   195
         Top             =   120
         Width           =   2700
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8760
         Top             =   360
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
         Caption         =   "Ado2"
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
         Left            =   10500
         Top             =   90
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
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   11775
      Begin VB.Frame frComments 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         TabIndex        =   184
         Top             =   5280
         Visible         =   0   'False
         Width           =   10935
         Begin VB.TextBox txtComments 
            Appearance      =   0  'Flat
            DataField       =   "EC_COMMENTS"
            Enabled         =   0   'False
            Height          =   1590
            Left            =   120
            MaxLength       =   600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   185
            Tag             =   "00-General Comments"
            Top             =   375
            Width           =   10605
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   186
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Frame frAccident 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   240
         TabIndex        =   140
         Top             =   -720
         Visible         =   0   'False
         Width           =   10935
         Begin VB.Frame frAccidentIllness 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   145
            Top             =   0
            Width           =   10815
            Begin VB.TextBox txtEmpPremises 
               Appearance      =   0  'Flat
               DataField       =   "EC_EMP_PREMISES"
               Enabled         =   0   'False
               Height          =   285
               Left            =   7200
               MaxLength       =   50
               TabIndex        =   148
               Tag             =   "01-Area where incident occurred"
               Top             =   315
               Width           =   3435
            End
            Begin VB.OptionButton optEmpPremises 
               Caption         =   "No"
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   147
               Tag             =   "40-Employee's Premises No"
               Top             =   315
               Width           =   615
            End
            Begin VB.OptionButton optEmpPremises 
               Caption         =   "Yes"
               Height          =   285
               Index           =   0
               Left            =   4560
               TabIndex        =   146
               Tag             =   "40-Employee's Premises Yes"
               Top             =   315
               Width           =   615
            End
            Begin INFOHR_Controls.CodeLookup clpCode 
               DataField       =   "EC_AREA"
               Height          =   285
               Index           =   13
               Left            =   7200
               TabIndex        =   149
               TabStop         =   0   'False
               Tag             =   "01-Area where incident occurred"
               Top             =   315
               Visible         =   0   'False
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   503
               ShowUnassigned  =   1
               TABLName        =   "ECPA"
               Enabled         =   0   'False
            End
            Begin VB.Label lblEmpPremises 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "EC_PREMISES"
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
               Height          =   165
               Left            =   4200
               TabIndex        =   187
               Top             =   480
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.Label lblSpecifyWhere 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Specify where"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6000
               TabIndex        =   170
               Top             =   360
               Width           =   1005
            End
            Begin VB.Label Label12 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Did the accident/illness happen on the employer's premises (owned, leased or maintained)?"
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   169
               Top             =   240
               Width           =   4275
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   171
            Top             =   600
            Width           =   10815
            Begin VB.TextBox txtOutsideProv 
               Appearance      =   0  'Flat
               DataField       =   "EC_OUTSIDE_CITY"
               Enabled         =   0   'False
               Height          =   285
               Left            =   7200
               MaxLength       =   55
               TabIndex        =   152
               Tag             =   "01-Location of the Incident"
               Top             =   315
               Width           =   3435
            End
            Begin VB.OptionButton optOutsideProvYN 
               Caption         =   "Yes"
               Height          =   285
               Index           =   0
               Left            =   4560
               TabIndex        =   150
               Tag             =   "40-Outside Province Yes"
               Top             =   315
               Width           =   615
            End
            Begin VB.OptionButton optOutsideProvYN 
               Caption         =   "No"
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   151
               Tag             =   "40-Outside Province No"
               Top             =   315
               Width           =   615
            End
            Begin VB.Label lblOutsideProvYN 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "EC_OUTSIDE_PROV"
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
               Height          =   165
               Left            =   4200
               TabIndex        =   188
               Top             =   480
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.Label Label14 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Did the accident/illness happen outside the Province of Ontario?"
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   173
               Top             =   240
               Width           =   4275
            End
            Begin VB.Label Label13 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Specify where"
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   6000
               TabIndex        =   172
               Top             =   337
               Width           =   1005
            End
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   120
            TabIndex        =   141
            Top             =   1200
            Width           =   10815
            Begin VB.TextBox txtWitness2 
               Appearance      =   0  'Flat
               DataField       =   "EC_WITNESS2"
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               MaxLength       =   100
               TabIndex        =   158
               Tag             =   "01-Current Position and Telephone #"
               Top             =   1080
               Width           =   6075
            End
            Begin VB.OptionButton optWitnessYN 
               Caption         =   "No"
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   154
               Tag             =   "40-Witness No"
               Top             =   315
               Width           =   615
            End
            Begin VB.OptionButton optWitnessYN 
               Caption         =   "Yes"
               Height          =   285
               Index           =   0
               Left            =   4560
               TabIndex        =   153
               Tag             =   "40-Witness Yes"
               Top             =   315
               Width           =   615
            End
            Begin VB.TextBox txtWitness1 
               Appearance      =   0  'Flat
               DataField       =   "EC_WITNESS1"
               Enabled         =   0   'False
               Height          =   285
               Left            =   4560
               MaxLength       =   100
               TabIndex        =   156
               Tag             =   "01-Current Position and Telephone #"
               Top             =   720
               Width           =   6075
            End
            Begin INFOHR_Controls.EmployeeLookup elpWitness1 
               DataField       =   "EC_WITNESS1_EMPNBR"
               Height          =   285
               Left            =   360
               TabIndex        =   155
               Tag             =   "01-Enter Employee Number"
               Top             =   720
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   503
               ShowUnassigned  =   1
               ShowDescription =   0   'False
               RefreshDescriptionWhen=   2
               Enabled         =   0   'False
            End
            Begin INFOHR_Controls.EmployeeLookup elpWitness2 
               DataField       =   "EC_WITNESS2_EMPNBR"
               Height          =   285
               Left            =   360
               TabIndex        =   157
               Tag             =   "01-Enter Employee Number"
               Top             =   1080
               Width           =   4005
               _ExtentX        =   7064
               _ExtentY        =   503
               ShowUnassigned  =   1
               ShowDescription =   0   'False
               RefreshDescriptionWhen=   2
               Enabled         =   0   'False
            End
            Begin VB.Label Label19 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "2."
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   144
               Top             =   1125
               Width           =   135
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "1."
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   143
               Top             =   765
               Width           =   135
            End
            Begin VB.Label Label15 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Are you aware of any witnesses or other employees involved in this accident/illness?"
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   142
               Top             =   240
               Width           =   4275
            End
            Begin VB.Label lblWitnessYN 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "EC_WITNESS"
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
               Height          =   165
               Left            =   4200
               TabIndex        =   189
               Top             =   480
               Visible         =   0   'False
               Width           =   285
            End
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   174
            Top             =   2640
            Width           =   10815
            Begin VB.TextBox txtResponsible1 
               Appearance      =   0  'Flat
               DataField       =   "EC_INDIV_NAME"
               Enabled         =   0   'False
               Height          =   285
               Left            =   880
               MaxLength       =   60
               TabIndex        =   161
               Tag             =   "01-Name and Work Phone #"
               Top             =   700
               Width           =   3435
            End
            Begin VB.OptionButton optResponsibleYN 
               Caption         =   "Yes"
               Height          =   285
               Index           =   0
               Left            =   4560
               TabIndex        =   159
               Tag             =   "40-Responsible Yes"
               Top             =   315
               Width           =   615
            End
            Begin VB.OptionButton optResponsibleYN 
               Caption         =   "No"
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   160
               Tag             =   "40-Responsible No"
               Top             =   315
               Width           =   615
            End
            Begin VB.TextBox txtResponsiblePhone 
               Appearance      =   0  'Flat
               DataField       =   "EC_INDIV_PHONE"
               Enabled         =   0   'False
               Height          =   285
               Left            =   6000
               MaxLength       =   40
               TabIndex        =   162
               Tag             =   "01-Name and Work Phone #"
               Top             =   700
               Width           =   2595
            End
            Begin VB.Label lblResponsibleYN 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "EC_INDIV_RESP"
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
               Height          =   165
               Left            =   4200
               TabIndex        =   190
               Top             =   480
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.Label Label20 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Work Phone #"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   4800
               TabIndex        =   177
               Top             =   750
               Width           =   1050
            End
            Begin VB.Label Label18 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Was any individual, who does not work for your firm, partially or totally responsible for this accident/illness?"
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   176
               Top             =   240
               Width           =   4275
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   120
               TabIndex        =   175
               Top             =   750
               Width           =   420
            End
         End
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   178
            Top             =   3720
            Width           =   10815
            Begin VB.CommandButton cmdBrowse 
               Caption         =   "..."
               Enabled         =   0   'False
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
               Left            =   10380
               TabIndex        =   166
               Tag             =   "Click to select the location"
               Top             =   302
               Width           =   375
            End
            Begin VB.OptionButton optPriorInjuryYN 
               Caption         =   "No"
               Height          =   285
               Index           =   1
               Left            =   5280
               TabIndex        =   164
               Tag             =   "40-Prior Incident No"
               Top             =   315
               Width           =   615
            End
            Begin VB.OptionButton optPriorInjuryYN 
               Caption         =   "Yes"
               Height          =   285
               Index           =   0
               Left            =   4560
               TabIndex        =   163
               Tag             =   "40-Prior Incident Yes"
               Top             =   315
               Width           =   615
            End
            Begin VB.TextBox txtPriorIncDate 
               Appearance      =   0  'Flat
               DataField       =   "EC_SIMILAR_INJ_DEATAILS"
               Enabled         =   0   'False
               Height          =   285
               Left            =   7740
               MaxLength       =   50
               TabIndex        =   165
               Tag             =   "01-Incident and Claim No."
               Top             =   305
               Width           =   2595
            End
            Begin VB.Label lblPriorInjuryYN 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "EC_SIMILAR_INJ"
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
               Height          =   165
               Left            =   4200
               TabIndex        =   191
               Top             =   480
               Visible         =   0   'False
               Width           =   285
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Are you aware of any prior similar or related problem, injury or condition?"
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   180
               Top             =   240
               Width           =   4275
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "Incident Date - Claim #"
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   6000
               TabIndex        =   179
               Top             =   345
               Width           =   1620
            End
         End
         Begin VB.Frame Frame8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   181
            Top             =   4320
            Width           =   10815
            Begin VB.CommandButton cmdImport 
               Caption         =   "Import"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   9600
               TabIndex        =   168
               Top             =   292
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox chkSubmissionAttch 
               Caption         =   "Submission Attached"
               DataField       =   "EC_ANY_CONCERNS"
               Height          =   195
               Left            =   4560
               TabIndex        =   167
               Tag             =   "Submit Concernes"
               Top             =   360
               Width           =   1965
            End
            Begin VB.Label lblImport 
               Alignment       =   1  'Right Justify
               Caption         =   "Submission"
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
               Height          =   195
               Left            =   7320
               TabIndex        =   183
               Top             =   360
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Image imgNoSec 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   9240
               Picture         =   "fehsinjrWF7.frx":6374
               Top             =   337
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Label Label24 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "If you have concerns about this claim, attach a written submission to this form."
               ForeColor       =   &H80000008&
               Height          =   435
               Left            =   120
               TabIndex        =   182
               Top             =   240
               Width           =   4275
            End
            Begin VB.Image imgSec 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   9240
               Picture         =   "fehsinjrWF7.frx":64BE
               Top             =   337
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin Threed.SSPanel frmDetails 
         Height          =   3135
         Left            =   240
         TabIndex        =   111
         Top             =   2880
         Width           =   10995
         _Version        =   65536
         _ExtentX        =   19394
         _ExtentY        =   5530
         _StockProps     =   15
         ForeColor       =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   0
         BorderWidth     =   0
         BevelOuter      =   0
         Begin VB.TextBox txtTask 
            Appearance      =   0  'Flat
            DataField       =   "EC_TASK"
            Height          =   285
            Left            =   2340
            MaxLength       =   40
            TabIndex        =   123
            Tag             =   "01-Task being performed when injured"
            Top             =   780
            Width           =   6050
         End
         Begin VB.TextBox txtOSHA300 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2340
            MaxLength       =   50
            TabIndex        =   113
            Tag             =   "00-Form 7/OSHA 300"
            Top             =   2760
            Visible         =   0   'False
            Width           =   6075
         End
         Begin VB.TextBox txtOSHACOM 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2340
            MaxLength       =   50
            TabIndex        =   112
            Tag             =   "00-Form 7 sec 6/ OSHA Comment"
            Top             =   2160
            Visible         =   0   'False
            Width           =   6075
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_JBCODE"
            Height          =   285
            Index           =   12
            Left            =   7200
            TabIndex        =   114
            Tag             =   "01-Position code"
            Top             =   1770
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECJB"
            MaxLength       =   6
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_LOC"
            Height          =   285
            Index           =   8
            Left            =   7200
            TabIndex        =   115
            Tag             =   "00-Location of Incident"
            Top             =   1440
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "EDLC"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_SECONDARY"
            Height          =   285
            Index           =   7
            Left            =   7200
            TabIndex        =   116
            Tag             =   "00-Secondary Cause of Injury"
            Top             =   1110
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECCA"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_SFACT"
            Height          =   285
            Index           =   10
            Left            =   7200
            TabIndex        =   117
            Tag             =   "00-Enter Facet - Code"
            Top             =   450
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECFA"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_SCODE"
            Height          =   285
            Index           =   11
            Left            =   7200
            TabIndex        =   118
            Tag             =   "00-Injury - Code"
            Top             =   120
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECCD"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_EQUIP"
            Height          =   285
            Index           =   6
            Left            =   2025
            TabIndex        =   119
            Tag             =   "00-Equipment being used when injured"
            Top             =   1770
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECEQ"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_CAUSECD"
            Height          =   285
            Index           =   4
            Left            =   2025
            TabIndex        =   120
            Tag             =   "01-Primary Cause of Injury"
            Top             =   1110
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECCA"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PFACT"
            Height          =   285
            Index           =   9
            Left            =   2025
            TabIndex        =   121
            Tag             =   "01-Enter Facet - Code"
            Top             =   450
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECFA"
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_CODE"
            Height          =   285
            Index           =   1
            Left            =   2025
            TabIndex        =   122
            Tag             =   "01-Injury - Code"
            Top             =   120
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECCD"
         End
         Begin VB.Label lblSecCause 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Secondary Cause"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   139
            Top             =   1155
            Width           =   1500
         End
         Begin VB.Label lblType 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Primary Injury"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   138
            Top             =   165
            Width           =   930
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Facet"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   137
            Top             =   495
            Width           =   405
         End
         Begin VB.Label lblTask 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Task"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   136
            Top             =   825
            Width           =   2205
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblCause 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Primary Cause"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   135
            Top             =   1155
            Width           =   1335
         End
         Begin VB.Label lblEquipment 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Equipment"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   134
            Top             =   1815
            Width           =   900
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Secondary Injury"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   133
            Top             =   165
            Width           =   1185
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Facet"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5520
            TabIndex        =   132
            Top             =   495
            Width           =   405
         End
         Begin VB.Label lblLocation 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   5520
            TabIndex        =   131
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label lblPosTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Job Code"
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
            Left            =   5520
            TabIndex        =   130
            Top             =   1815
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label lblUpdateDate 
            Caption         =   "Updated Date"
            Height          =   255
            Left            =   3960
            TabIndex        =   129
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label lblUpdDateDesc 
            Caption         =   "lblUserDesc"
            Height          =   255
            Left            =   5040
            TabIndex        =   128
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Label lblUpdateBy 
            Caption         =   "Updated By"
            Height          =   255
            Left            =   90
            TabIndex        =   127
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lblUserDesc 
            Caption         =   "lblUserDesc"
            Height          =   255
            Left            =   1080
            TabIndex        =   126
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label lblOSHA300 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Form 7/OSHA 300"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   125
            Top             =   2805
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label lblOSHACOM 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Form 7 sec 6/ OSHA Comment"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   90
            TabIndex        =   124
            Top             =   2205
            Visible         =   0   'False
            Width           =   2355
         End
      End
      Begin VB.CommandButton cmdPageLeft 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   9720
         Picture         =   "fehsinjrWF7.frx":6608
         Style           =   1  'Graphical
         TabIndex        =   193
         Tag             =   "Grant All Basic"
         Top             =   0
         Width           =   720
      End
      Begin VB.CommandButton cmdPageRight 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   10560
         Picture         =   "fehsinjrWF7.frx":6A4A
         Style           =   1  'Graphical
         TabIndex        =   192
         Tag             =   "Grant All Basic"
         Top             =   0
         Width           =   705
      End
      Begin VB.CheckBox chkCompleted 
         Caption         =   "Check1"
         DataField       =   "EC_INJURED_ONLINE"
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
         Left            =   3360
         TabIndex        =   21
         Tag             =   "Completed"
         Top             =   150
         Width           =   195
      End
      Begin VB.OptionButton OptInjDis 
         Caption         =   "Disease"
         Height          =   285
         Index           =   1
         Left            =   1260
         TabIndex        =   16
         Tag             =   "40-Disease"
         Top             =   120
         Width           =   1035
      End
      Begin VB.OptionButton OptInjDis 
         Caption         =   " Other"
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   15
         Tag             =   "40-Other"
         Top             =   120
         Width           =   1035
      End
      Begin VB.OptionButton OptInjDis 
         Caption         =   "Injury"
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Tag             =   "40-Injury"
         Top             =   120
         Width           =   855
      End
      Begin VB.Frame frInjuries 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   10935
         Begin VB.CheckBox chkHead 
            Caption         =   "Head"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   780
         End
         Begin VB.CheckBox chkEars 
            Caption         =   "Ear(s)"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   1200
            Width           =   780
         End
         Begin VB.CheckBox chkTeeth 
            Caption         =   "Teeth"
            Height          =   195
            Left            =   1200
            TabIndex        =   60
            Top             =   480
            Width           =   765
         End
         Begin VB.CheckBox chkNeck 
            Caption         =   "Neck"
            Height          =   195
            Left            =   1200
            TabIndex        =   59
            Top             =   720
            Width           =   765
         End
         Begin VB.CheckBox chkChest 
            Caption         =   "Chest"
            Height          =   195
            Left            =   1200
            TabIndex        =   58
            Top             =   960
            Width           =   765
         End
         Begin VB.CheckBox chkUpBack 
            Caption         =   "Upper Back"
            Height          =   195
            Left            =   2280
            TabIndex        =   57
            Top             =   480
            Width           =   1215
         End
         Begin VB.CheckBox chkLowBack 
            Caption         =   "Lower Back"
            Height          =   195
            Left            =   2280
            TabIndex        =   56
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkAbdomen 
            Caption         =   "Abdomen"
            Height          =   195
            Left            =   2280
            TabIndex        =   55
            Top             =   960
            Width           =   1065
         End
         Begin VB.CheckBox chkPelvis 
            Caption         =   "Pelvis"
            Height          =   195
            Left            =   2280
            TabIndex        =   54
            Top             =   1200
            Width           =   825
         End
         Begin VB.CheckBox chkLShoulder 
            Caption         =   " Shoulder"
            Height          =   195
            Left            =   3720
            TabIndex        =   53
            Top             =   480
            Width           =   1035
         End
         Begin VB.CheckBox chkLArm 
            Caption         =   "    Arm"
            Height          =   195
            Left            =   3720
            TabIndex        =   52
            Top             =   720
            Width           =   885
         End
         Begin VB.CheckBox chkLElbow 
            Caption         =   "  Elbow"
            Height          =   195
            Left            =   3720
            TabIndex        =   51
            Top             =   960
            Width           =   885
         End
         Begin VB.CheckBox chkLForearm 
            Caption         =   " Forearm"
            Height          =   195
            Left            =   3720
            TabIndex        =   50
            Top             =   1200
            Width           =   1005
         End
         Begin VB.CheckBox chkRShoulder 
            Height          =   195
            Left            =   4800
            TabIndex        =   49
            Top             =   480
            Width           =   350
         End
         Begin VB.CheckBox chkRArm 
            Height          =   195
            Left            =   4800
            TabIndex        =   48
            Top             =   720
            Width           =   350
         End
         Begin VB.CheckBox chkRForearm 
            Height          =   195
            Left            =   4800
            TabIndex        =   46
            Top             =   1200
            Width           =   350
         End
         Begin VB.CheckBox chkLHip 
            Caption         =   "     Hip"
            Height          =   195
            Left            =   7440
            TabIndex        =   45
            Top             =   480
            Width           =   1020
         End
         Begin VB.CheckBox chkLThigh 
            Caption         =   "   Thigh"
            Height          =   195
            Left            =   7440
            TabIndex        =   44
            Top             =   720
            Width           =   1020
         End
         Begin VB.CheckBox chkLKnee 
            Caption         =   "    Knee"
            Height          =   195
            Left            =   7440
            TabIndex        =   43
            Top             =   960
            Width           =   1020
         End
         Begin VB.CheckBox chkLLowLeg 
            Caption         =   " Lower Leg"
            Height          =   195
            Left            =   7440
            TabIndex        =   42
            Top             =   1200
            Width           =   1155
         End
         Begin VB.CheckBox chkRHip 
            Height          =   195
            Left            =   8640
            TabIndex        =   41
            Top             =   480
            Width           =   350
         End
         Begin VB.CheckBox chkRThigh 
            Height          =   195
            Left            =   8640
            TabIndex        =   40
            Top             =   720
            Width           =   350
         End
         Begin VB.CheckBox chkRKnee 
            Height          =   195
            Left            =   8640
            TabIndex        =   39
            Top             =   960
            Width           =   350
         End
         Begin VB.CheckBox chkRLowLeg 
            Height          =   195
            Left            =   8640
            TabIndex        =   38
            Top             =   1200
            Width           =   350
         End
         Begin VB.CheckBox chkLWrist 
            Caption         =   "   Wrist"
            Height          =   195
            Left            =   5520
            TabIndex        =   37
            Top             =   480
            Width           =   885
         End
         Begin VB.CheckBox chkLHand 
            Caption         =   "   Hand"
            Height          =   195
            Left            =   5520
            TabIndex        =   36
            Top             =   720
            Width           =   885
         End
         Begin VB.CheckBox chkLFingers 
            Caption         =   " Finger(s)"
            Height          =   195
            Left            =   5520
            TabIndex        =   35
            Top             =   960
            Width           =   1005
         End
         Begin VB.CheckBox chkRWrist 
            Height          =   195
            Left            =   6600
            TabIndex        =   34
            Top             =   480
            Width           =   350
         End
         Begin VB.CheckBox chkRHand 
            Height          =   195
            Left            =   6600
            TabIndex        =   33
            Top             =   720
            Width           =   350
         End
         Begin VB.CheckBox chkRFingers 
            Height          =   195
            Left            =   6600
            TabIndex        =   32
            Top             =   960
            Width           =   350
         End
         Begin VB.CheckBox chkLAnkle 
            Caption         =   "  Ankle"
            Height          =   195
            Left            =   9480
            TabIndex        =   31
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkLFoot 
            Caption         =   "   Foot"
            Height          =   195
            Left            =   9480
            TabIndex        =   30
            Top             =   720
            Width           =   825
         End
         Begin VB.CheckBox chkLToes 
            Caption         =   " Toe(s)"
            Height          =   195
            Left            =   9480
            TabIndex        =   29
            Top             =   960
            Width           =   900
         End
         Begin VB.CheckBox chkRAnkle 
            Height          =   195
            Left            =   10440
            TabIndex        =   28
            Top             =   480
            Width           =   350
         End
         Begin VB.CheckBox chkRToes 
            Height          =   195
            Left            =   10440
            TabIndex        =   26
            Top             =   960
            Width           =   350
         End
         Begin VB.CheckBox chkOther 
            Caption         =   "Other"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1605
            Width           =   765
         End
         Begin VB.CheckBox chkEyes 
            Caption         =   "Eye(s)"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   825
         End
         Begin VB.CheckBox chkFace 
            Caption         =   "Face"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   780
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_OTHER"
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   63
            Tag             =   "01-Enter Body Site - Code"
            Top             =   1560
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECBS"
            Enabled         =   0   'False
         End
         Begin VB.CheckBox chkRElbow 
            Height          =   195
            Left            =   4800
            TabIndex        =   47
            Top             =   960
            Width           =   350
         End
         Begin VB.CheckBox chkRFoot 
            Height          =   195
            Left            =   10440
            TabIndex        =   27
            Top             =   720
            Width           =   350
         End
         Begin INFOHR_Controls.CodeLookup clpCode 
            DataField       =   "EC_PBODY"
            Height          =   285
            Index           =   2
            Left            =   5640
            TabIndex        =   194
            Tag             =   "01-Enter Body Site - Code"
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            ShowUnassigned  =   1
            TABLName        =   "ECBS"
         End
         Begin VB.Line Line4 
            X1              =   9240
            X2              =   9240
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Line Line3 
            X1              =   7200
            X2              =   7200
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Line Line2 
            X1              =   3600
            X2              =   3600
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Line Line1 
            X1              =   5280
            X2              =   5280
            Y1              =   240
            Y2              =   1440
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            Left            =   3675
            TabIndex        =   110
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Left            =   4680
            TabIndex        =   109
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            Left            =   5475
            TabIndex        =   108
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            Left            =   7380
            TabIndex        =   107
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Left"
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
            Left            =   9435
            TabIndex        =   106
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Left            =   6480
            TabIndex        =   105
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Left            =   8520
            TabIndex        =   104
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Right"
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
            Left            =   10320
            TabIndex        =   103
            Top             =   240
            Width           =   525
         End
         Begin VB.Label lblHead 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_HEAD"
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
            Left            =   120
            TabIndex        =   102
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblFace 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_FACE"
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
            TabIndex        =   101
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblEyes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_EYES"
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
            Left            =   360
            TabIndex        =   100
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblEars 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_EARS"
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
            Left            =   480
            TabIndex        =   99
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblTeeth 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_TEETH"
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
            Left            =   1200
            TabIndex        =   98
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblChest 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_CHEST"
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
            Left            =   1560
            TabIndex        =   96
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblUpBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_UPPER_BACK"
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
            Left            =   2400
            TabIndex        =   95
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLoBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LOWER_BACK"
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
            Left            =   2520
            TabIndex        =   94
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblAbdomen 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_ABDOMEN"
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
            Left            =   2640
            TabIndex        =   93
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblPelvis 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_PELVIS"
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
            Left            =   2760
            TabIndex        =   92
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLShoulder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_SHOULDER"
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
            Left            =   3720
            TabIndex        =   91
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLArm 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_ARM"
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
            Left            =   3840
            TabIndex        =   90
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLElbow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_ELBOW"
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
            Left            =   3960
            TabIndex        =   89
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLForerm 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_FOREARM"
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
            Left            =   4080
            TabIndex        =   88
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRShoulder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_SHOULDER"
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
            Left            =   4320
            TabIndex        =   87
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRArm 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_ARM"
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
            Left            =   4440
            TabIndex        =   86
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRElbow 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_ELBOW"
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
            Left            =   4560
            TabIndex        =   85
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRForearm 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_FOREARM"
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
            Left            =   4680
            TabIndex        =   84
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLWrist 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_WRIST"
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
            Left            =   5640
            TabIndex        =   83
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLHand 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_HAND"
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
            Left            =   5760
            TabIndex        =   82
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLFingers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_FINGER"
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
            TabIndex        =   81
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRHand 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_HAND"
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
            Left            =   6240
            TabIndex        =   79
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRFingers 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_FINGER"
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
            Left            =   6360
            TabIndex        =   78
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLHip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_HIP"
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
            Left            =   7680
            TabIndex        =   77
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLLoLeg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_LOWER_LEG"
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
            Left            =   8040
            TabIndex        =   74
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRHip 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_HIP"
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
            Left            =   8280
            TabIndex        =   73
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRThigh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_THIGH"
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
            Left            =   8400
            TabIndex        =   72
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRKnee 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_KNEE"
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
            Left            =   8520
            TabIndex        =   71
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRLoLeg 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_LOWER_LEG"
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
            Left            =   8640
            TabIndex        =   70
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLAnkle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_ANKLE"
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
            Left            =   9600
            TabIndex        =   69
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLFoot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_FOOT"
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
            Left            =   9720
            TabIndex        =   68
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLToes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_TOES"
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
            Left            =   9840
            TabIndex        =   67
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRFoot 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_FOOT"
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
            Left            =   10200
            TabIndex        =   65
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRToes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_TOES"
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
            Left            =   10320
            TabIndex        =   64
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblNeck 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_NECK"
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
            Left            =   1320
            TabIndex        =   97
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRWrist 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_WRIST"
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
            Left            =   6120
            TabIndex        =   80
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLThigh 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_THIGH"
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
            Left            =   7800
            TabIndex        =   76
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblLKnee 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_LFT_KNEE"
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
            Left            =   7920
            TabIndex        =   75
            Top             =   1440
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Label lblRAnkle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "EC_RGT_ANKLE"
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
            Left            =   10080
            TabIndex        =   66
            Top             =   1200
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin MSComctlLib.TabStrip tbAccidentInjury 
         Height          =   5775
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   10186
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Injury Details"
               Key             =   "keyInjury"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Other Injury Details"
               Key             =   "keyOther"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Comments"
               Key             =   "keyComments"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblIncident 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number"
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
         Left            =   6120
         TabIndex        =   20
         Top             =   150
         Width           =   1410
      End
      Begin VB.Label lblIncidentNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         DataField       =   "EC_CASE"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8160
         TabIndex        =   19
         Top             =   150
         Width           =   90
      End
      Begin VB.Label lblSINJ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "EC_SINJ"
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
         Height          =   285
         Left            =   4860
         TabIndex        =   18
         Top             =   105
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Injured on Line"
         Height          =   195
         Index           =   1
         Left            =   3630
         TabIndex        =   17
         Top             =   150
         Width           =   1050
      End
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EC_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1650
      TabIndex        =   8
      Top             =   10230
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "001"
      DataField       =   "EC_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   390
      TabIndex        =   9
      Top             =   10230
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmEHSINJURYWF7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbUpd%
Dim fglbNew
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim glbOccDate
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsDATA1 As New ADODB.Recordset


Private Function chkHSInjury()
Dim SQLQ As String, Msg As String, dd#, X%

chkHSInjury = False

On Error GoTo chkHSInjury_Err

'If Len(clpCode(1).Text) < 1 Then
'    If Not OptInjDis(2).Value Then
'        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
'        MsgBox "Injury Code is a required field"
'        clpCode(1).SetFocus
'        Exit Function
'    End If
'Else
    If clpCode(1).Caption = "Unassigned" Then
        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
        MsgBox "Injury code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
'End If

If chkOther.Value = 1 And Len(Trim(clpCode(0).Text)) = 0 Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
    MsgBox "'Other' area of injury is selected, 'Other' code is required"
    clpCode(0).SetFocus
    Exit Function
End If

If clpCode(0).Caption = "Unassigned" Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
    If Not clpCode(0).ListChecker Then Exit Function
End If

'WF7
'If Len(clpCode(2).Text) < 1 Then
'    If Not OptInjDis(2).Value Then
'        MsgBox "#1 Body Site is a required field"
'        clpCode(2).SetFocus
'        Exit Function
'    End If
'Else
'    If clpCode(2).Caption = "Unassigned" Then
'        MsgBox "#1 Body Site code must be valid"
'        clpCode(2).SetFocus
'        Exit Function
'    End If
'End If

'WF7
'If Len(clpCode(3).Text) > 0 Then
'  If clpCode(3).Caption = "Unassigned" Then
'    MsgBox "#2 Body Site code must be valid"
'    clpCode(3).SetFocus
'    Exit Function
'  End If
'  ' dkostka - 10/02/2001 - Removed on request of Linda

''  If clpCode(3) = clpCode(2) Then
''    MsgBox "#2 Body Site can not be the same as #1 Body Site."
''    clpCode(3).SetFocus
''    Exit Function
''  End If
'End If

If clpCode(9).Caption = "Unassigned" And Not OptInjDis(2).Value Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
    MsgBox "Invalid Facet Code"
    clpCode(9).SetFocus
    Exit Function
End If

'Jerry asked this to not be mandatory - Form 7
'If Not glbWFC Then
'    If Len(txtTask) < 1 And Not OptInjDis(2).Value Then
'        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
'        MsgBox "Task is a required field"
'        txtTask.SetFocus
'        Exit Function
'    End If
'End If

If Len(clpCode(4).Text) < 1 Then
    If Not OptInjDis(2).Value Then
        If glbWFC Then
            tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
            MsgBox "Pers Eq #1 is a required field"
            clpCode(4).SetFocus
            Exit Function
        'Else       'As per Next Release Documentation
            'MsgBox "Primary cause is a required field"
        End If
        'MsgBox "Primary cause is a required field"
    End If
Else
    If clpCode(4).Caption = "Unassigned" Then
        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
        If glbWFC Then
            MsgBox "Pers Eq #1 must be valid"
        Else
            MsgBox "Cause code must be valid"
        End If
        clpCode(4).SetFocus
        Exit Function
    End If
End If

'WF7
'If Len(clpCode(5).Text) < 1 Then
'    If Not OptInjDis(2).Value Then
'        'MsgBox "Plant Area code is a required field"   'As per Next Release Documentation
'        'clpCode(5).SetFocus
'        'Exit Function
'    End If
'    If glbWFC Then
'        MsgBox "Plant Area is a required field"
'        clpCode(5).SetFocus
'        Exit Function
'    End If
'Else
'    If Len(clpCode(5).Text) > 1 Then
'        If clpCode(5).Caption = "Unassigned" Then
'            MsgBox "Plant Area code must be valid"
'            clpCode(5).SetFocus
'            Exit Function
'        End If
'    End If
'End If


For X% = 6 To 12
    If X% <> 12 Then
        If Len(clpCode(X%).Text) > 0 Then
            If clpCode(X%).Caption = "Unassigned" Then
                tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
                MsgBox "Invalid code Entered"
                clpCode(X%).SetFocus
                Exit Function
            End If
        End If
    End If
Next

If glbWFC Then
    If Len(clpCode(6)) = 0 Then
        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
        MsgBox "Equipment is a required field"
        clpCode(6).SetFocus
        Exit Function
    End If
    If Len(Trim(clpCode(12).Text)) < 1 Then
        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
        MsgBox "Job Code is a required field"
        clpCode(12).SetFocus
        Exit Function
    Else
        If clpCode(12).Caption = "Unassigned" Then
            tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
            MsgBox "Job Code must be valid"
            clpCode(12).SetFocus
            Exit Function
        End If
    End If
End If

'If clpCode(1) = clpCode(11) Then
'  MsgBox "#2 Injury Code can not be the same as #1 Injury Code."
'  clpCode(11).SetFocus
'  Exit Function
'End If
'~~~~~~~~~~~~~~commented out by RAUBREY 6/2/97 ~~~~~~~~~~~~~~
'If clpCode(9) = clpCode(10) And Len(clpCode(9)) > 0 Then
'  MsgBox "#2 Facet can not be the same as #1 Facet."
'  clpCode(10).SetFocus
'  Exit Function
'End If

If glbLinamar Then
    'If Len(clpCode(4)) = 0 Then
    '    MsgBox lblCause.Caption & " is a required field"
    '    clpCode(4).SetFocus
    '    Exit Function
    'End If
    If Len(clpCode(5)) = 0 Then
        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
        MsgBox "Plant Area is a required field"
        clpCode(5).SetFocus
        Exit Function
    End If
    If Len(clpCode(6)) = 0 Then
        tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(1)
        MsgBox "Equipment is a required field"
        clpCode(6).SetFocus
        Exit Function
    End If
    'Ticket #14666
    'If Len(clpCode(11)) = 0 Then
    '    MsgBox "Secondary Injury is a required field"
    '    clpCode(11).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(10)) = 0 Then
    '    MsgBox "Facet is a required field"
    '    clpCode(10).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(3)) = 0 Then
    '    MsgBox "Body Site is a required field"
    '    clpCode(3).SetFocus
    '    Exit Function
    'End If
    'If Len(clpCode(8)) = 0 Then
    '    MsgBox lblLocation(1).Caption & " is a required field"
    '    clpCode(8).SetFocus
    '    Exit Function
    'End If
    'If Len(txtComments.Text) = 0 Then  'Ticket #16782
    '    MsgBox "Comments is a required field"
    '    txtComments.SetFocus
    '    Exit Function
    'End If
End If

'Specify Where - can be specified even when No is selected.
''If optEmpPremises(0).Value And Len(Trim(clpCode(13).Text)) = 0 Then
'If optEmpPremises(0).Value And Len(Trim(txtEmpPremises.Text)) = 0 Then
'    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
'    MsgBox "If Accident/Illness happened on the employer's premises then 'Specify where' cannot be blank.", vbExclamation
'    txtEmpPremises.SetFocus
'    Exit Function
''ElseIf optEmpPremises(1).Value And Len(Trim(clpCode(13).Text)) > 0 Then
'ElseIf optEmpPremises(1).Value And Len(Trim(txtEmpPremises.Text)) > 0 Then
'    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
'    MsgBox "If Accident/Illness DID NOT happen on the employer's premises then 'Specify where' should be blank.", vbExclamation
'    txtEmpPremises.SetFocus
'    Exit Function
'End If
'If clpCode(13).Caption = "Unassigned" Then
'    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
'    If Not clpCode(13).ListChecker Then Exit Function
'End If

If optOutsideProvYN(0).Value And Len(Trim(txtOutsideProv.Text)) = 0 Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If Accident/Illness happened outside the Province of Ontario then 'Specify where' cannot be blank.", vbExclamation
    txtOutsideProv.SetFocus
    Exit Function
ElseIf optOutsideProvYN(1).Value And Len(Trim(txtOutsideProv.Text)) > 0 Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If Accident/Illness DID NOT happen outside the Province of Ontario then 'Specify where' should be blank.", vbExclamation
    txtOutsideProv.SetFocus
    Exit Function
End If

If optWitnessYN(0).Value And (Len(Trim(elpWitness1.Text)) = 0 And Len(txtWitness1.Text) = 0) Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If you are aware of any witnesses or other employees involved in this accident/illness then specify their Employee # or Name/Position/Work Phone # in '1.' and/or '2.'", vbExclamation
    elpWitness1.SetFocus
    Exit Function
ElseIf optWitnessYN(1).Value And (Len(Trim(elpWitness1.Text)) > 0 Or Len(Trim(elpWitness2.Text)) > 0 Or Len(txtWitness1.Text) > 0 Or Len(txtWitness2.Text) > 0) Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If you are NOT aware of any witnesses or other employees involved in this accident/illness then Employee #(s) or Name/Position/Work Phone # in '1.' and '2.' should be blank.", vbExclamation
    elpWitness1.SetFocus
    Exit Function
End If

'Ticket #22368 - Phone # not mandatory
'If optResponsibleYN(0).Value And (Len(Trim(txtResponsible1.Text)) = 0 Or Len(Trim(txtResponsiblePhone.Text)) = 0) Then
If optResponsibleYN(0).Value And (Len(Trim(txtResponsible1.Text)) = 0) Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    'MsgBox "If any individual, who does not work for your firm, was partially or totally responsible for this accident/illness then 'Name' and 'Work Phone' cannot be blank.", vbExclamation
    MsgBox "If any individual, who does not work for your firm, was partially or totally responsible for this accident/illness then 'Name' cannot be blank.", vbExclamation
    txtResponsible1.SetFocus
    Exit Function
ElseIf optResponsibleYN(1).Value And (Len(Trim(txtResponsible1.Text)) > 0 Or Len(Trim(txtResponsiblePhone.Text)) > 0) Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If NO individual, who does not work for your firm, was partially or totally responsible for this accident/illness then 'Name' and 'Work Phone' should be blank.", vbExclamation
    optResponsibleYN(1).SetFocus
    Exit Function
End If

If optPriorInjuryYN(0).Value And Len(Trim(txtPriorIncDate.Text)) = 0 Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If you are aware of any prior similar or related problem, injury or condition then 'Incident Date - Claim #' cannot be blank.", vbExclamation
    txtPriorIncDate.SetFocus
    Exit Function
ElseIf optPriorInjuryYN(1).Value And Len(Trim(txtPriorIncDate.Text)) > 0 Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "If you are NOT aware of any prior similar or related problem, injury or condition then 'Incident Date - Claim #' should be blank.", vbExclamation
    txtPriorIncDate.SetFocus
    Exit Function
End If

If chkSubmissionAttch.Value = 1 And imgNoSec.Visible = True Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "'Submission Attached' is selected, but not written submission attached.", vbExclamation
    'chkSubmissionAttch.SetFocus
    'Exit Function
ElseIf chkSubmissionAttch.Value <> 1 And imgNoSec.Visible = False And gsAttachment_DB = True Then
    tbAccidentInjury.SelectedItem = tbAccidentInjury.Tabs(2)
    MsgBox "'Submission Attached' not is selected, but written submission is attached.", vbExclamation
    chkSubmissionAttch.SetFocus
    Exit Function
End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If OptInjDis(0) Then lblSINJ = "I"
If OptInjDis(1) Then lblSINJ = "D"
If OptInjDis(2) Then lblSINJ = "O"

chkHSInjury = True

Exit Function

chkHSInjury_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkInjury", "HR_OCC_HEALTH_SAFETY", "edit/Add")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

'Private Sub cmdCAction_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdCancel_Click()
Dim locDate, X
On Error GoTo Can_Err

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh

''' Sam add July 2002 * Remove Binding Control
If Not (rsDATA.EOF And rsDATA.BOF) Then rsDATA.CancelUpdate

'Form 7 - Additional Sections - Section F
If Not (rsDATA1.EOF And rsDATA1.BOF) Then rsDATA1.CancelUpdate

Call Display_Value

fglbNew = False

'Call ST_UPD_MODE(True)  ' reset screen's attributes
Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OCC_HEALTH_SAFETY", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()

Call NextForm

Unload Me

If glbOnTop = "FRMEHSINJURYWF7" Then glbOnTop = ""

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

'Call ST_UPD_MODE(True)

Call SET_UP_MODE

'clpCode(1).SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_OCC_HEALTH_SAFETY", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Function cmdOK_Click()
Dim X
Dim xBPart As String
Dim SQLQ As String
Dim xSat As Integer
Dim xSunDate, xSatDate

On Error GoTo Add_Err

cmdOK_Click = False

If Not chkHSInjury() Then Exit Function

'Option button's data field value setting
If optEmpPremises(0) Then lblEmpPremises = "1" Else lblEmpPremises = "0"
If optOutsideProvYN(0) Then lblOutsideProvYN = "1" Else lblOutsideProvYN = "0"
If optWitnessYN(0) Then lblWitnessYN = "1" Else lblWitnessYN = "0"
If optResponsibleYN(0) Then lblResponsibleYN = "1" Else lblResponsibleYN = "0"
If optPriorInjuryYN(0) Then lblPriorInjuryYN = "1" Else lblPriorInjuryYN = "0"

'Update the Body Site field. Jerry said to update with the first Body Parts checkbox found selected
'This is done so the Body Site report works.
xBPart = ""
xBPart = First_Body_Part_Selected
If Len(xBPart) > 0 Then
    clpCode(2).Text = xBPart
Else
    clpCode(2).Text = ""
End If

rsDATA.Requery

'Ticket #28288 Franks 03/07/2016
If rsDATA.EOF And rsDATA.BOF Then
    Exit Function
End If

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If

'Form 7 - Additional Sections - Section F
If fglbNew Then
    Call Additional_Form7Sections_Update
End If
'    If Not Data1.Recordset.EOF Then
'        If rsDATA1.State <> 0 Then: If rsDATA1.EOF Then rsDATA1.Close Else If rsDATA1.EditMode = adEditAdd Then rsDATA1.CancelUpdate: rsDATA1.Close Else rsDATA1.Close
'        If glbtermopen Then
'            SQLQ = "SELECT " & FldList1
'            SQLQ = SQLQ & " FROM Term_OHS_FORM7_SECTIONS"
'            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'            SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
'            rsDATA1.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'            Data2.RecordSource = SQLQ
'        Else
'            SQLQ = "SELECT " & FldList1
'            SQLQ = SQLQ & " FROM HR_OHS_FORM7_SECTIONS"
'            SQLQ = SQLQ & " WHERE F7_EMPNBR = " & glbLEE_ID
'            SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
'            rsDATA1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            Data2.RecordSource = SQLQ
'        End If
'        If rsDATA1.EOF Then
'            rsDATA1.AddNew
'            rsDATA1("F7_COMPNO") = "001"
'            rsDATA1("F7_EMPNBR") = glbLEE_ID
'            rsDATA1("F7_CASE") = Data1.Recordset!EC_CASE
'            rsDATA1("F7_FED_AMT") = GetEmpData(glbLEE_ID, "ED_TD1DOL", "")
'            rsDATA1("F7_PROV_AMT") = GetEmpData(glbLEE_ID, "ED_PROVAMT", "")
'            If glbtermopen Then
'                rsDATA1("TERM_SEQ") = glbTERM_Seq
'            End If
'
'            'Week 1 - 4 From/To Dates - for Section H-8 and I-C
'            'Week is Sun - Sat.
'            If IsDate(Data1.Recordset!EC_OCCDATE) Then
'
'                'Get which day of the week it is and get the # of days before the last Saturday
'                Select Case Weekday(Data1.Recordset!EC_OCCDATE)
'                    Case vbMonday
'                        xSat = -2
'                    Case vbTuesday
'                        xSat = -3
'                    Case vbWednesday
'                        xSat = -4
'                    Case vbThursday
'                        xSat = -5
'                    Case vbFriday
'                        xSat = -6
'                    Case vbSaturday
'                        xSat = -7
'                    Case vbSunday
'                        xSat = -8
'                End Select
'
'                'Week 1 - From
'                'Compute the date of the last Saturday
'                xSatDate = DateAdd("d", xSat, Data1.Recordset!EC_OCCDATE)
'
'                'Compute the date of last Sunday - 1 week prior
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK1") = xSunDate
'                rsDATA1("F7_FWEEK1") = xSunDate
'
'                'Week 1 - To
'                rsDATA1("F7_OTH_EARN_TO_WK1") = xSatDate
'                rsDATA1("F7_TWEEK1") = xSatDate
'
'
'                'Week 2 - From
'                xSatDate = DateAdd("d", -1, xSunDate)
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK2") = xSunDate
'                rsDATA1("F7_FWEEK2") = xSunDate
'
'                'Week 2 - To
'                rsDATA1("F7_OTH_EARN_TO_WK2") = xSatDate
'                rsDATA1("F7_TWEEK2") = xSatDate
'
'
'                'Week 3 - From
'                xSatDate = DateAdd("d", -1, xSunDate)
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK3") = xSunDate
'                rsDATA1("F7_FWEEK3") = xSunDate
'
'                'Week 3 - To
'                rsDATA1("F7_OTH_EARN_TO_WK3") = xSatDate
'                rsDATA1("F7_TWEEK3") = xSatDate
'
'
'                'Week 4 - From
'                xSatDate = DateAdd("d", -1, xSunDate)
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK4") = xSunDate
'                rsDATA1("F7_FWEEK4") = xSunDate
'
'                'Week 4 - To
'                rsDATA1("F7_OTH_EARN_TO_WK4") = xSatDate
'                rsDATA1("F7_TWEEK4") = xSatDate
'            End If
'
'            rsDATA1.Update
'        End If
'    End If
'End If


Data1.Refresh

fglbNew = False

'If gsAttachment_DB Then
'    If glbDocNewRecord Then 'New Record only
'        If Len(glbDocImpFile) > 0 Then
'            glbDocKey = xID
'            glbJob = rsDATA("DE_CASE")
'            glbDocTmp = rsDATA("DE_DOCNO")
'            Call AttachmentAdd(glbLEE_ID, glbDocImpFile)
'        End If
'    End If
'    glbDocImpFile = ""
'End If

cmdOK_Click = True

'Call ST_UPD_MODE(True)

Call SET_UP_MODE

X = NextFormIF("Injury/Location")


Exit Function

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OCC_HEALTH_SAFETY", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Injury"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Injury"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdTCause_Click()
'frmEHSCause.Show
'Unload Me
'End Sub

'Private Sub cmdTCause_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub

'Private Sub cmdWCBMed_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Function EERetrieve()
Dim SQLQ As String
Dim xSat As Integer
Dim xSunDate, xSatDate

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERError

If glbtermopen Then
    SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
Else
    SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Data1.RecordSource = SQLQ
Data1.Refresh


'Form 7 - Additional Sections
Call Additional_Form7Sections_Update
'If Not Data1.Recordset.EOF Then
'    If rsDATA1.State <> 0 Then: If rsDATA1.EOF Then rsDATA1.Close Else If rsDATA1.EditMode = adEditAdd Then rsDATA1.CancelUpdate: rsDATA1.Close Else rsDATA1.Close
'    If glbtermopen Then
'        SQLQ = "SELECT " & FldList1
'        SQLQ = SQLQ & " FROM Term_OHS_FORM7_SECTIONS"
'        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'        SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
'        rsDATA1.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'        Data2.RecordSource = SQLQ
'    Else
'        SQLQ = "SELECT " & FldList1
'        SQLQ = SQLQ & " FROM HR_OHS_FORM7_SECTIONS"
'        SQLQ = SQLQ & " WHERE F7_EMPNBR = " & glbLEE_ID
'        SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
'        rsDATA1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        Data2.RecordSource = SQLQ
'    End If
'    If rsDATA1.EOF Then
'        rsDATA1.AddNew
'        rsDATA1("F7_COMPNO") = "001"
'        rsDATA1("F7_EMPNBR") = glbLEE_ID
'        rsDATA1("F7_CASE") = Data1.Recordset!EC_CASE
'        rsDATA1("F7_FED_AMT") = GetEmpData(glbLEE_ID, "ED_TD1DOL", "")
'        rsDATA1("F7_PROV_AMT") = GetEmpData(glbLEE_ID, "ED_PROVAMT", "")
'        If glbtermopen Then
'            rsDATA1("TERM_SEQ") = glbTERM_Seq
'        End If
'
'        'Week 1 - 4 From/To Dates - for Section H-8 and I-C
'        'Week is Sun - Sat.
'        If IsDate(Data1.Recordset!EC_OCCDATE) Then
'
'            'Get which day of the week it is and get the # of days before the last Saturday
'            Select Case Weekday(Data1.Recordset!EC_OCCDATE)
'                Case vbMonday
'                    xSat = -2
'                Case vbTuesday
'                    xSat = -3
'                Case vbWednesday
'                    xSat = -4
'                Case vbThursday
'                    xSat = -5
'                Case vbFriday
'                    xSat = -6
'                Case vbSaturday
'                    xSat = -7
'                Case vbSunday
'                    xSat = -8
'            End Select
'
'            'Week 1 - From
'            'Compute the date of the last Saturday
'            xSatDate = DateAdd("d", xSat, Data1.Recordset!EC_OCCDATE)
'
'            'Compute the date of last Sunday - 1 week prior
'            xSunDate = DateAdd("d", -6, xSatDate)
'            rsDATA1("F7_OTH_EARN_FROM_WK1") = xSunDate
'            rsDATA1("F7_FWEEK1") = xSunDate
'
'            'Week 1 - To
'            rsDATA1("F7_OTH_EARN_TO_WK1") = xSatDate
'            rsDATA1("F7_TWEEK1") = xSatDate
'
'
'            'Week 2 - From
'            xSatDate = DateAdd("d", -1, xSunDate)
'            xSunDate = DateAdd("d", -6, xSatDate)
'            rsDATA1("F7_OTH_EARN_FROM_WK2") = xSunDate
'            rsDATA1("F7_FWEEK2") = xSunDate
'
'            'Week 2 - To
'            rsDATA1("F7_OTH_EARN_TO_WK2") = xSatDate
'            rsDATA1("F7_TWEEK2") = xSatDate
'
'
'            'Week 3 - From
'            xSatDate = DateAdd("d", -1, xSunDate)
'            xSunDate = DateAdd("d", -6, xSatDate)
'            rsDATA1("F7_OTH_EARN_FROM_WK3") = xSunDate
'            rsDATA1("F7_FWEEK3") = xSunDate
'
'            'Week 3 - To
'            rsDATA1("F7_OTH_EARN_TO_WK3") = xSatDate
'            rsDATA1("F7_TWEEK3") = xSatDate
'
'
'            'Week 4 - From
'            xSatDate = DateAdd("d", -1, xSunDate)
'            xSunDate = DateAdd("d", -6, xSatDate)
'            rsDATA1("F7_OTH_EARN_FROM_WK4") = xSunDate
'            rsDATA1("F7_FWEEK4") = xSunDate
'
'            'Week 4 - To
'            rsDATA1("F7_OTH_EARN_TO_WK4") = xSatDate
'            rsDATA1("F7_TWEEK4") = xSatDate
'        End If
'
'        rsDATA1.Update
'    End If
'End If

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OCC_HEALTH_SAFETY", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function

End Function

Private Sub chkAbdomen_Click()
    If chkAbdomen.Value = 1 Then
        lblAbdomen = "ABMN"
    Else
        lblAbdomen = ""
    End If
End Sub

Private Sub chkChest_Click()
    If chkChest.Value = 1 Then
        lblChest = "CHST"
    Else
        lblChest = ""
    End If
End Sub

Private Sub chkEars_Click()
    If chkEars.Value = 1 Then
        lblEars = "EAR"
    Else
        lblEars = ""
    End If
End Sub

Private Sub chkEyes_Click()
    If chkEyes.Value = 1 Then
        lblEyes = "EYE"
    Else
        lblEyes = ""
    End If
End Sub

Private Sub chkFace_Click()
    If chkFace.Value = 1 Then
        lblFace = "FACE"
    Else
        lblFace = ""
    End If
End Sub

Private Sub chkHead_Click()
    If chkHead.Value = 1 Then
        lblHead = "HEAD"
    Else
        lblHead = ""
    End If
End Sub

Private Sub chkLAnkle_Click()
    If chkLAnkle.Value = 1 Then
        lblLAnkle = "ANKL"
    Else
        lblLAnkle = ""
    End If
End Sub

Private Sub chkLArm_Click()
    If chkLArm.Value = 1 Then
        lblLArm = "ARML"
    Else
        lblLArm = ""
    End If
End Sub

Private Sub chkLElbow_Click()
    If chkLElbow.Value = 1 Then
        lblLElbow = "ELBL"
    Else
        lblLElbow = ""
    End If
End Sub

Private Sub chkLFingers_Click()
    If chkLFingers.Value = 1 Then
        lblLFingers = "FNGL"
    Else
        lblLFingers = ""
    End If
End Sub

Private Sub chkLFoot_Click()
    If chkLFoot.Value = 1 Then
        lblLFoot = "FOTL"
    Else
        lblLFoot = ""
    End If
End Sub

Private Sub chkLForearm_Click()
    If chkLForearm.Value = 1 Then
        lblLForerm = "FAML"
    Else
        lblLForerm = ""
    End If
End Sub

Private Sub chkLHand_Click()
    If chkLHand.Value = 1 Then
        lblLHand = "HNDL"
    Else
        lblLHand = ""
    End If
End Sub

Private Sub chkLHip_Click()
    If chkLHip.Value = 1 Then
        lblLHip = "HIPL"
    Else
        lblLHip = ""
    End If
End Sub

Private Sub chkLKnee_Click()
    If chkLKnee.Value = 1 Then
        lblLKnee = "KNEL"
    Else
        lblLKnee = ""
    End If
End Sub

Private Sub chkLLowLeg_Click()
    If chkLLowLeg.Value = 1 Then
        lblLLoLeg = "LLGL"
    Else
        lblLLoLeg = ""
    End If
End Sub

Private Sub chkLowBack_Click()
    If chkLowBack.Value = 1 Then
        lblLoBack = "LBCK"
    Else
        lblLoBack = ""
    End If
End Sub

Private Sub chkLShoulder_Click()
    If chkLShoulder.Value = 1 Then
        lblLShoulder = "SHDL"
    Else
        lblLShoulder = ""
    End If
End Sub

Private Sub chkLThigh_Click()
    If chkLThigh.Value = 1 Then
        lblLThigh = "THGL"
    Else
        lblLThigh = ""
    End If
End Sub

Private Sub chkLToes_Click()
    If chkLToes.Value = 1 Then
        lblLToes = "TOEL"
    Else
        lblLToes = ""
    End If
End Sub

Private Sub chkLWrist_Click()
    If chkLWrist.Value = 1 Then
        lblLWrist = "WRTL"
    Else
        lblLWrist = ""
    End If
End Sub

Private Sub chkNeck_Click()
    If chkNeck.Value = 1 Then
        lblNeck = "NECK"
    Else
        lblNeck = ""
    End If
End Sub

Private Sub chkOther_Click()
    If chkOther.Value = 1 Then
        clpCode(0).Enabled = True
    Else
        clpCode(0).Text = ""
        clpCode(0).Enabled = False
    End If
End Sub

Private Sub chkPelvis_Click()
    If chkPelvis.Value = 1 Then
        lblPelvis = "PLVS"
    Else
        lblPelvis = ""
    End If
End Sub

Private Sub chkRAnkle_Click()
    If chkRAnkle.Value = 1 Then
        lblRAnkle = "ANKR"
    Else
        lblRAnkle = ""
    End If
End Sub

Private Sub chkRArm_Click()
    If chkRArm.Value = 1 Then
        lblRArm = "ARMR"
    Else
        lblRArm = ""
    End If
End Sub

Private Sub chkRElbow_Click()
    If chkRElbow.Value = 1 Then
        lblRElbow = "ELBR"
    Else
        lblRElbow = ""
    End If
End Sub

Private Sub chkRFingers_Click()
    If chkRFingers.Value = 1 Then
        lblRFingers = "FNGR"
    Else
        lblRFingers = ""
    End If
End Sub

Private Sub chkRFoot_Click()
    If chkRFoot.Value = 1 Then
        lblRFoot = "FOTR"
    Else
        lblRFoot = ""
    End If
End Sub

Private Sub chkRForearm_Click()
    If chkRForearm.Value = 1 Then
        lblRForearm = "FAMR"
    Else
        lblRForearm = ""
    End If
End Sub

Private Sub chkRHand_Click()
    If chkRHand.Value = 1 Then
        lblRHand = "HNDR"
    Else
        lblRHand = ""
    End If
End Sub

Private Sub chkRHip_Click()
    If chkRHip.Value = 1 Then
        lblRHip = "HIPR"
    Else
        lblRHip = ""
    End If
End Sub

Private Sub chkRKnee_Click()
    If chkRKnee.Value = 1 Then
        lblRKnee = "KNER"
    Else
        lblRKnee = ""
    End If
End Sub

Private Sub chkRLowLeg_Click()
    If chkRLowLeg.Value = 1 Then
        lblRLoLeg = "LLGR"
    Else
        lblRLoLeg = ""
    End If
End Sub

Private Sub chkRShoulder_Click()
    If chkRShoulder.Value = 1 Then
        lblRShoulder = "SHDR"
    Else
        lblRShoulder = ""
    End If
End Sub

Private Sub chkRThigh_Click()
    If chkRThigh.Value = 1 Then
        lblRThigh = "THGR"
    Else
        lblRThigh = ""
    End If
End Sub

Private Sub chkRToes_Click()
    If chkRToes.Value = 1 Then
        lblRToes = "TOER"
    Else
        lblRToes = ""
    End If
End Sub

Private Sub chkRWrist_Click()
    If chkRWrist.Value = 1 Then
        lblRWrist = "WRTR"
    Else
        lblRWrist = ""
    End If
End Sub

Private Sub chkSubmissionAttch_Click()
    If chkSubmissionAttch.Value = 1 Then
        cmdImport.Enabled = True
    Else
        cmdImport.Enabled = False
    End If
End Sub

Private Sub chkTeeth_Click()
    If chkTeeth.Value = 1 Then
        lblTeeth = "TETH"
    Else
        lblTeeth = ""
    End If
End Sub

Private Sub chkUpBack_Click()
    If chkUpBack.Value = 1 Then
        lblUpBack = "UBCK"
    Else
        lblUpBack = ""
    End If
End Sub

Private Sub clpCode_Change(Index As Integer)
'    If Len(clpCode(13)) > 0 Then
'        optEmpPremises(0).Value = True
'    End If
    
    If Index = 0 Then
        If Len(clpCode(0)) > 0 Then
            chkOther.Value = 1
        Else
            chkOther.Value = 0
        End If
    
    End If
End Sub

Private Sub cmdBrowse_Click()
    frmIncidentList.Show 1
    DoEvents
End Sub

Private Sub cmdF7Sections_Click()
    frmEInjF7Sections.txtIncidentNo.Text = Me.lblIncidentNo.Caption
    frmEInjF7Sections.txtIncidentDate.Text = Data1.Recordset("EC_OCCDATE")
    frmEInjF7Sections.Show 1
End Sub

Private Sub cmdImport_Click()
Dim xID
    glbDocNewRecord = False
    glbDocName = "INJURYWF7"
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        glbDocKey = 0
        glbJob = ""
        glbDocTmp = ""
    Else
        glbDocKey = rsDATA("EC_ID")
        glbJob = rsDATA("EC_CASE")
        'glbDocTmp = rsDATA("EC_DOCKEY")
    End If

    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEHSINJURYWF7")
End Sub

Private Sub cmdPageLeft_Click(Index As Integer)
    'Save the data
    If Not cmdOK_Click() Then Exit Sub
    
    'Unload the current form and load the next one
    Unload Me
    
    'Next form
    Screen.MousePointer = HOURGLASS
    Load frmEHSINCIDENT
    frmEHSINCIDENT.ZOrder 0
    Screen.MousePointer = DEFAULT

End Sub

Private Sub cmdPageRight_Click(Index As Integer)
    'Save the data
    If Not cmdOK_Click() Then Exit Sub
    
    'Unload the current form and load the next one
    Unload Me
    
    'Next form
    Screen.MousePointer = HOURGLASS
    Load frmEHSWCB
    frmEHSWCB.ZOrder 0
    Screen.MousePointer = DEFAULT

End Sub

Private Sub elpWitness1_LostFocus()
    Dim xJob, xPhone
    'Get emmployee's Current Position and Telephone #2
    If Len(elpWitness1) > 0 Then
        If elpWitness1.ListChecker Then
            'Call function to retrieve Phone #
            xPhone = GetEmpData(elpWitness1, "ED_BUSNBR", "")
            
            'Call function to retrieve Current Position
            xJob = GetJHData(elpWitness1, "JH_JOB", "")
            If Len(xJob) > 0 Then
                xJob = GetJobData(xJob, "JB_DESCR")
            End If
            'txtWitness1.Text = "Position: " & xJob & "; Phone: " & Format(xPhone, "(###) ###-####")
            txtWitness1.Text = "" & xJob '& "; " & Format(xPhone, "(###) ###-####")
            If Len(xPhone) > 0 Then
                txtWitness1.Text = txtWitness1.Text & "; " & Format(xPhone, "(###) ###-####")
            End If
        End If
    End If
End Sub

Private Sub elpWitness2_LostFocus()
    Dim xJob, xPhone
    'Get emmployee's Current Position and Telephone #2
    If Len(elpWitness2) > 0 Then
        If elpWitness2.ListChecker Then
            'Call function to retrieve Phone #
            xPhone = GetEmpData(elpWitness2, "ED_BUSNBR", "")
            
            'Call function to retrieve Current Position
            xJob = GetJHData(elpWitness2, "JH_JOB", "")
            If Len(xJob) > 0 Then
                xJob = GetJobData(xJob, "JB_DESCR")
            End If
            'txtWitness2.Text = "Position: " & xJob & "; Phone: " & Format(xPhone, "(###) ###-####")
            txtWitness2.Text = "" & xJob '& "; " & Format(xPhone, "(###) ###-####")
            If Len(xPhone) > 0 Then
                txtWitness2.Text = txtWitness2.Text & "; " & Format(xPhone, "(###) ###-####")
            End If
        End If
    End If
End Sub

'Private Sub cmdWSIB_Click()
'frmEHSWCBC.Show
'Unload Me
'End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
' Me.cmdModify_Click
glbOnTop = "FRMEHSINJURYWF7"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEHSINJURYWF7"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer, X% ' records found
glbOnTop = "FRMEHSINJURYWF7"
fglbUpd% = False

'Data1.DatabaseName = glbIHRDB
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data2.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data2.ConnectionString = glbAdoIHRDB
End If
Screen.MousePointer = DEFAULT

lblLocation(1).Caption = lStr(lblLocation(1).Caption)

If glbWFC Then
    lblCause = "Pers Eq #1"
    lblSecCause = "Pers Eq #2"
    clpCode(4).TABLTitle = "Personal Equipment 1"
    clpCode(4).Tag = "01-Personal Equipment 1"
    clpCode(7).TABLTitle = "Personal Equipment 2"
    clpCode(7).Tag = "01-Personal Equipment 2"
    vbxTrueGrid.Columns(6).Caption = "Pers Eq #1"
    vbxTrueGrid.Columns(7).Caption = "Pers Eq #2"
    lblPosTitle.Visible = True
    clpCode(12).Visible = True
    lblTask.Font.Bold = False
    lblCause.FontBold = True
    'lblPlantArea.FontBold = True 'WF7
    lblEquipment.FontBold = True
End If

If glbCompSerial = "S/N - 2387W" Then  'Bird Packaging Limited 'Ticket #13636
    OptInjDis(0).Caption = "Sudden"
    OptInjDis(2).Caption = "Gradual"
    lblTitle(1).Caption = " Other"
    lblType.Caption = "Accident Type"
    Label1.Caption = "Area of Injury"
    Label2.Caption = "Area of Injury"
    'lblBody2.Caption = "N/A"       'WF7
    lblCause.Caption = "Injury Class #1"
    lblSecCause.Caption = "Injury Class #2"
End If

'Ticket #14573
If glbLinamar Then
    'lblCause.Caption = "Root Cause"
    lblCause.Visible = False
    'clpCode(4).Tag = "01-Root Cause of Injury"
    clpCode(4).Visible = False
    lblSecCause.Visible = False
    clpCode(7).Visible = False
    
    'Ticket #14703
    lblLocation(1).Visible = False
    clpCode(8).Visible = False
    
    'Mandatory fields
    'lblCause.FontBold = True
    'lblPlantArea.FontBold = True   'WF7
    lblEquipment.FontBold = True
    'Label3.FontBold = True     'Ticket #14666
    'Label1.FontBold = True     'Ticket #14666
    'lblBody2.FontBold = True   'Ticket #14666
    'lblLocation(1).FontBold = True 'Ticket #14666
    'lblTitle(14).FontBold = True   'Ticket #16782
    
    'Ticket #15172
    OptInjDis(0).Visible = False
    OptInjDis(1).Visible = False
    OptInjDis(2).Visible = False
    chkCompleted.Visible = False
    lblTitle(1).Visible = False
    'lblOSHA300.Top = lblCause.Top   'Ticket #15172
    'txtOSHA300.Top = clpCode(4).Top 'Ticket #15172
    lblOSHA300.Visible = False  'Ticket #15172
    
    'Hemu
    lblTask.Caption = "Task (Form 7/OSHA 300)"
    lblOSHACOM.Top = lblCause.Top '2325
    txtOSHACOM.Top = clpCode(4).Top '2280
    'lblPlantArea.Top = 2325 '2685  'WF7
    clpCode(5).Top = 2280   '2640
    
    lblEquipment.Top = 2685 '3030
    clpCode(6).Top = 2640   '3000
    lblTitle(14).Top = 3030 '3430
    txtComments.Top = 3280  '3680
    'Hemu
    
    lblOSHACOM.Visible = True
    txtOSHA300.Visible = False  'Ticket #15172
    txtOSHACOM.Visible = True
    txtOSHA300.DataField = "EC_OSHA300"
    txtOSHACOM.DataField = "EC_OSHACOM"
End If

If glbWSIBModule Then
    tbAccidentInjury.Tabs(2).Caption = "Accident/Illness Dates and Details (Continued)"
    tbAccidentInjury.Tabs(3).Caption = "Describe what happened to cause injury..."
    lblTitle(14).Visible = False
End If

If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(False)
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(False)
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(Str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.SetFocus

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Injury - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Injury - " & Left$(glbLEE_SName, 5)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call Display_Value
'Call ST_UPD_MODE(False)

If Not gSec_Upd_HSW7Injury And glbWSIBModule Then
 '   cmdModify.Enabled = False
End If

Call INI_Controls(Me)
 
If glbLinamar Then
    'lblPlantArea.FontBold = True   'WF7
End If

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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
Dim c As Long

On Error GoTo Eh

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 10500 Then
        scrControl.Value = 0
        
        Frame6.Top = 3000
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 4000
        
        scrControl.Max = 3500
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
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Form 7 Injury/Location Form", "Resize")
    Resume exH
End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."

Set frmEHSINJURYWF7 = Nothing 'carmen may 00

Call NextForm

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

fglbUpd% = YN

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If


fUPMode = TF    ' update mode
frmDetails.Enabled = TF
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
txtComments.Enabled = TF
'vbxTrueGrid.Enabled = FT

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT

'cmdPrint.Enabled = FT

'cmdWCBMed.Enabled = FT
'cmdIncident.Enabled = FT
'cmdCAction.Enabled = FT
'cmdContact.Enabled = FT
'cmdTCause.Enabled = FT
'cmdWSIB.Enabled = FT

OptInjDis(0).Enabled = TF
OptInjDis(1).Enabled = TF
OptInjDis(2).Enabled = TF
clpCode(1).Enabled = TF
'clpCode(2).Enabled = TF    'WF7
'clpCode(3).Enabled = TF    'WF7
clpCode(4).Enabled = TF
'clpCode(5).Enabled = TF    'WF7
clpCode(6).Enabled = TF
clpCode(7).Enabled = TF
clpCode(8).Enabled = TF
clpCode(9).Enabled = TF
clpCode(10).Enabled = TF
clpCode(11).Enabled = TF
clpCode(12).Enabled = TF
txtComments.Enabled = TF
txtTask.Enabled = TF

If Data1.Recordset.BOF And Data1.Recordset.EOF Then 'Add by Frank 8/21/2001
    'cmdModify.Enabled = False
    cmdF7Sections.Enabled = False
    chkSubmissionAttch.Enabled = False
Else
    'Me.cmdModify_Click
    cmdF7Sections.Enabled = True
    chkSubmissionAttch.Enabled = True
End If

glbJob = ""
glbSDate = "01/01/1900"
glbDocKey = 0
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    If Data1.Recordset("EC_ANY_CONCERNS") = True And Not IsNull(Data1.Recordset("EC_DOCKEY")) Then
        glbJob = Data1.Recordset("EC_CASE")
        glbSDate = Data1.Recordset("EC_OCCDATE")
        glbDocKey = Data1.Recordset("EC_DOCKEY")
        'glbDocTmp = IIf(IsNull(Data1.Recordset("EC_DOCKEY")), "", Data1.Recordset("EC_DOCKEY"))
    Else
        glbJob = ""
        glbSDate = ""
        glbDocKey = ""
        glbDocTmp = ""
    End If
End If

glbDocName = "INJURYWF7"
If gsAttachment_DB Then
    Call DispimgIcon(Me, "frmEHSINJURYWF7")
    If gSec_Upd_HSW7Injury And glbWSIBModule And Not glbtermopen Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
            If Data1.Recordset("EC_ANY_CONCERNS") = True Then
                cmdImport.Enabled = True
            Else
                cmdImport.Enabled = False
            End If
        End If
    End If
End If
 
'Updstats(3) = Data3.Recordset("EC_OCCDATE")
'Updstats(4) = glbDocName   'WF7

End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEHSINJURYWF7")
    Call FillMemoFile(SQLQ, "INJURYWF7")
End Sub

Private Sub lblAbdomen_Change()
    If Len(lblAbdomen) > 0 Then
        chkAbdomen.Value = 1
    Else
        chkAbdomen.Value = 0
    End If
End Sub

Private Sub lblChest_Change()
    If Len(lblChest) > 0 Then
        chkChest.Value = 1
    Else
        chkChest.Value = 0
    End If
End Sub

Private Sub lblEars_Change()
    If Len(lblEars) > 0 Then
        chkEars.Value = 1
    Else
        chkEars.Value = 0
    End If
End Sub

Private Sub lblEmpPremises_Change()
    If lblEmpPremises = "" Then
        optEmpPremises(0) = False
        optEmpPremises(1) = False
    ElseIf lblEmpPremises <> "0" Then
        optEmpPremises(0) = True
    Else
        optEmpPremises(1) = True
    End If
End Sub

Private Sub lblEyes_Change()
    If Len(lblEyes) > 0 Then
        chkEyes.Value = 1
    Else
        chkEyes.Value = 0
    End If
End Sub

Private Sub lblFace_Change()
    If Len(lblFace) > 0 Then
        chkFace.Value = 1
    Else
        chkFace.Value = 0
    End If
End Sub

Private Sub lblHead_Change()
    If Len(lblHead) > 0 Then
        chkHead.Value = 1
    Else
        chkHead.Value = 0
    End If
End Sub

Private Sub lblLAnkle_Change()
    If Len(lblLAnkle) > 0 Then
        chkLAnkle.Value = 1
    Else
        chkLAnkle.Value = 0
    End If
End Sub

Private Sub lblLArm_Change()
    If Len(lblLArm) > 0 Then
        chkLArm.Value = 1
    Else
        chkLArm.Value = 0
    End If
End Sub

Private Sub lblLElbow_Change()
    If Len(lblLElbow) > 0 Then
        chkLElbow.Value = 1
    Else
        chkLElbow.Value = 0
    End If
End Sub

Private Sub lblLFingers_Change()
    If Len(lblLFingers) > 0 Then
        chkLFingers.Value = 1
    Else
        chkLFingers.Value = 0
    End If
End Sub

Private Sub lblLFoot_Change()
    If Len(lblLFoot) > 0 Then
        chkLFoot.Value = 1
    Else
        chkLFoot.Value = 0
    End If
End Sub

Private Sub lblLForerm_Change()
    If Len(lblLForerm) > 0 Then
        chkLForearm.Value = 1
    Else
        chkLForearm.Value = 0
    End If
End Sub

Private Sub lblLHand_Change()
    If Len(lblLHand) > 0 Then
        chkLHand.Value = 1
    Else
        chkLHand.Value = 0
    End If
End Sub

Private Sub lblLHip_Change()
    If Len(lblLHip) > 0 Then
        chkLHip.Value = 1
    Else
        chkLHip.Value = 0
    End If
End Sub

Private Sub lblLKnee_Change()
    If Len(lblLKnee) > 0 Then
        chkLKnee.Value = 1
    Else
        chkLKnee.Value = 0
    End If
End Sub

Private Sub lblLLoLeg_Change()
    If Len(lblLLoLeg) > 0 Then
        chkLLowLeg.Value = 1
    Else
        chkLLowLeg.Value = 0
    End If
End Sub

Private Sub lblLoBack_Change()
    If Len(lblLoBack) > 0 Then
        chkLowBack.Value = 1
    Else
        chkLowBack.Value = 0
    End If
End Sub

Private Sub lblLShoulder_Change()
    If Len(lblLShoulder) > 0 Then
        chkLShoulder.Value = 1
    Else
        chkLShoulder.Value = 0
    End If
End Sub

Private Sub lblLThigh_Change()
    If Len(lblLThigh) > 0 Then
        chkLThigh.Value = 1
    Else
        chkLThigh.Value = 0
    End If
End Sub

Private Sub lblLToes_Change()
    If Len(lblLToes) > 0 Then
        chkLToes.Value = 1
    Else
        chkLToes.Value = 0
    End If
End Sub

Private Sub lblLWrist_Change()
    If Len(lblLWrist) > 0 Then
        chkLWrist.Value = 1
    Else
        chkLWrist.Value = 0
    End If
End Sub

Private Sub lblNeck_Change()
    If Len(lblNeck) > 0 Then
        chkNeck.Value = 1
    Else
        chkNeck.Value = 0
    End If
End Sub

Private Sub lblOutsideProvYN_Change()
    If lblOutsideProvYN = "" Then
        optOutsideProvYN(0) = False
        optOutsideProvYN(1) = False
    ElseIf lblOutsideProvYN <> "0" Then
        optOutsideProvYN(0) = True
    Else
        optOutsideProvYN(1) = True
    End If
End Sub

Private Sub lblPelvis_Change()
    If Len(lblPelvis) > 0 Then
        chkPelvis.Value = 1
    Else
        chkPelvis.Value = 0
    End If
End Sub

Private Sub lblPriorInjuryYN_Change()
    If lblPriorInjuryYN = "" Then
        optPriorInjuryYN(0) = False
        optPriorInjuryYN(1) = False
    ElseIf lblPriorInjuryYN <> "0" Then
        optPriorInjuryYN(0) = True
    Else
        optPriorInjuryYN(1) = True
    End If
End Sub

Private Sub lblRAnkle_Change()
    If Len(lblRAnkle) > 0 Then
        chkRAnkle.Value = 1
    Else
        chkRAnkle.Value = 0
    End If
End Sub

Private Sub lblRArm_Change()
    If Len(lblRArm) > 0 Then
        chkRArm.Value = 1
    Else
        chkRArm.Value = 0
    End If
End Sub

Private Sub lblRElbow_Change()
    If Len(lblRElbow) > 0 Then
        chkRElbow.Value = 1
    Else
        chkRElbow.Value = 0
    End If
End Sub

Private Sub lblResponsibleYN_Change()
    If lblResponsibleYN = "" Then
        optResponsibleYN(0) = False
        optResponsibleYN(1) = False
    ElseIf lblResponsibleYN <> "0" Then
        optResponsibleYN(0) = True
    Else
        optResponsibleYN(1) = True
    End If
End Sub

Private Sub lblRFingers_Change()
    If Len(lblRFingers) > 0 Then
        chkRFingers.Value = 1
    Else
        chkRFingers.Value = 0
    End If
End Sub

Private Sub lblRFoot_Change()
    If Len(lblRFoot) > 0 Then
        chkRFoot.Value = 1
    Else
        chkRFoot.Value = 0
    End If
End Sub

Private Sub lblRForearm_Change()
    If Len(lblRForearm) > 0 Then
        chkRForearm.Value = 1
    Else
        chkRForearm.Value = 0
    End If
End Sub

Private Sub lblRHand_Change()
    If Len(lblRHand) > 0 Then
        chkRHand.Value = 1
    Else
        chkRHand.Value = 0
    End If
End Sub

Private Sub lblRHip_Change()
    If Len(lblRHip) > 0 Then
        chkRHip.Value = 1
    Else
        chkRHip.Value = 0
    End If
End Sub

Private Sub lblRKnee_Change()
    If Len(lblRKnee) > 0 Then
        chkRKnee.Value = 1
    Else
        chkRKnee.Value = 0
    End If
End Sub

Private Sub lblRLoLeg_Change()
    If Len(lblRLoLeg) > 0 Then
        chkRLowLeg.Value = 1
    Else
        chkRLowLeg.Value = 0
    End If
End Sub

Private Sub lblRShoulder_Change()
    If Len(lblRShoulder) > 0 Then
        chkRShoulder.Value = 1
    Else
        chkRShoulder.Value = 0
    End If
End Sub

Private Sub lblRThigh_Change()
    If Len(lblRThigh) > 0 Then
        chkRThigh.Value = 1
    Else
        chkRThigh.Value = 0
    End If
End Sub

Private Sub lblRToes_Change()
    If Len(lblRToes) > 0 Then
        chkRToes.Value = 1
    Else
        chkRToes.Value = 0
    End If
End Sub

Private Sub lblRWrist_Change()
    If Len(lblRWrist) > 0 Then
        chkRWrist.Value = 1
    Else
        chkRWrist.Value = 0
    End If
End Sub

Private Sub lblSINJ_Change()
  If lblSINJ = "I" Then OptInjDis(0) = True
  If lblSINJ = "D" Then OptInjDis(1) = True
  If lblSINJ = "O" Then OptInjDis(2) = True
End Sub

Private Sub lblTeeth_Change()
    If Len(lblTeeth) > 0 Then
        chkTeeth.Value = 1
    Else
        chkTeeth.Value = 0
    End If
End Sub

Private Sub lblUpBack_Change()
    If Len(lblUpBack) > 0 Then
        chkUpBack.Value = 1
    Else
        chkUpBack.Value = 0
    End If
End Sub

Private Sub lblWitnessYN_Change()
    If lblWitnessYN = "" Then
        optWitnessYN(0) = False
        optWitnessYN(1) = False
    ElseIf lblWitnessYN <> "0" Then
        optWitnessYN(0) = True
    Else
        optWitnessYN(1) = True
    End If
End Sub

Private Sub optEmpPremises_Click(Index As Integer)
    If optEmpPremises(0).Value = True Then
        'clpCode(13).Enabled = True
        txtEmpPremises.Enabled = True
    Else
        'clpCode(13).Text = ""
        'clpCode(13).Enabled = False
        'txtEmpPremises.Text = ""
        txtEmpPremises.Enabled = True
    End If
End Sub

Private Sub optOutsideProvYN_Click(Index As Integer)
    If optOutsideProvYN(0).Value = True Then
        txtOutsideProv.Enabled = True
    Else
        txtOutsideProv.Text = ""
        txtOutsideProv.Enabled = False
    End If
End Sub

Private Sub OptInjDis_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optPriorInjuryYN_Click(Index As Integer)
    If optPriorInjuryYN(0).Value = True Then
        txtPriorIncDate.Enabled = True
        cmdBrowse.Enabled = True
    Else
        txtPriorIncDate.Text = ""
        txtPriorIncDate.Enabled = False
        cmdBrowse.Enabled = False
    End If
End Sub

Private Sub optResponsibleYN_Click(Index As Integer)
    If optResponsibleYN(0).Value = True Then
        txtResponsible1.Enabled = True
        txtResponsiblePhone.Enabled = True
    Else
        txtResponsible1.Text = ""
        txtResponsible1.Enabled = False
        txtResponsiblePhone.Enabled = False
    End If
End Sub

Private Sub optWitnessYN_Click(Index As Integer)
    If optWitnessYN(0).Value = True Then
        elpWitness1.Enabled = True
        elpWitness2.Enabled = True
        txtWitness1.Enabled = True
        txtWitness2.Enabled = True
    Else
        elpWitness1.Text = ""
        elpWitness2.Text = ""
        elpWitness1.Enabled = False
        elpWitness2.Enabled = False
        txtWitness1.Text = ""
        txtWitness2.Text = ""
        txtWitness1.Enabled = False
        txtWitness2.Enabled = False
    End If
End Sub

Private Sub scrControl_Change()
Frame6.Top = 3000 - scrControl.Value
End Sub

Private Sub tbAccidentInjury_Click()
    If tbAccidentInjury.SelectedItem.Index = 1 Then
        frAccident.Visible = False
        frComments.Visible = False
        'frELostTime.Visible = False
        'frFReturnToWork.Visible = False
        frInjuries.Visible = True
        frmDetails.Visible = True
    ElseIf tbAccidentInjury.SelectedItem.Index = 2 Then
        frInjuries.Visible = False
        frmDetails.Visible = False
        frComments.Visible = False
        'frELostTime.Visible = False
        'frFReturnToWork.Visible = False
        frAccident.Top = 960
        frAccident.Visible = True
    ElseIf tbAccidentInjury.SelectedItem.Index = 3 Then
        frInjuries.Visible = False
        frmDetails.Visible = False
        frAccident.Visible = False
        'frELostTime.Visible = False
        'frFReturnToWork.Visible = False
        frComments.Top = 960
        frComments.Visible = True
    ElseIf tbAccidentInjury.SelectedItem.Index = 4 Then
        frInjuries.Visible = False
        frmDetails.Visible = False
        frAccident.Visible = False
        frComments.Visible = False
        'frFReturnToWork.Visible = False
        'frELostTime.Top = 960
        'frELostTime.Visible = True
    ElseIf tbAccidentInjury.SelectedItem.Index = 5 Then
        frInjuries.Visible = False
        frmDetails.Visible = False
        frAccident.Visible = False
        frComments.Visible = False
        'frELostTime.Visible = False
        'frFReturnToWork.Top = 960
        'frFReturnToWork.Visible = True
    End If
    
End Sub

'Private Sub clpCode_DblClick(Index As Integer)
'Dim oCode As String, OCodeD As String
'oCode = clpCode(Index)
'OCodeD = clpCode(Index)
'Call Get_Code(CodeCodes(Index, 1), CodeCodes(Index, 2))
'If glbCodeRef Then Call ReCreatSnap(Index)
'If Len(glbCode) < 1 Then
'    clpCode(Index).Text = oCode
'    clpCode(Index).Caption = OCodeD
'Else
'    clpCode(Index).Text = glbCode
'    clpCode(Index).Caption = glbCodeDesc
'    clpCode(Index).Visible = True
'End If
'End Sub
'Private Sub clpCode_GotFocus(Index As Integer)
'clpCode(1).Tag = "01-Injury - Code"
'clpCode(2).Tag = "01-Enter Body Site - Code"
'If glbWFC Then
'  clpCode(4).Tag = "01-Pers Eq #1 of Injury"
'  clpCode(7).Tag = "01-Pers Eq #2 of Injury"
'Else
'  clpCode(4).Tag = "01-Primary Cause of Injury"
'  clpCode(7).Tag = "01-Secondary Cause of Injury"
'End If
'clpCode(5).Tag = "01-Area in plant/building where occurred"
'clpCode(9).Tag = "01-Enter Facet - Code"
'clpCode(11).Tag = "01-Injury - Code"
'Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub clpCode_KeyPress(Index As Integer, KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

Private Sub txtComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEmpPremises_Change()
    If Len(txtEmpPremises.Text) > 0 Then
        'optEmpPremises(0).Value = True
    End If
End Sub

Private Sub txtEmpPremises_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtOSHA300_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtOSHACOM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtOutsideProv_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTask_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
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
    
    If glbtermopen Then
        SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & "WHERE TERM_SEQ=" & glbTERM_Seq
    Else
        SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & "WHERE EC_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
    
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    'Set FRS = Data1.Recordset.Clone
    'vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim X%
Call Display_Value
End Sub

Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "EC_EMPNBR, EC_CASE, EC_OCCDATE, EC_CODE, "
SQLQ = SQLQ & "EC_PBODY, EC_SBODY, EC_CAUSECD, EC_AREA, EC_EQUIP,"
SQLQ = SQLQ & "EC_TASK, EC_LOC, EC_SECONDARY, EC_COMPNO,"
SQLQ = SQLQ & "EC_COMMENTS, EC_PFACT, EC_SFACT, EC_SCODE, EC_SINJ,"
SQLQ = SQLQ & "EC_INJURED_ONLINE,"
SQLQ = SQLQ & "EC_JBCODE, EC_LDATE , EC_LTIME, EC_LUSER, EC_ID, EC_DOCKEY,"
SQLQ = SQLQ & "EC_HEAD,EC_FACE,EC_EYES,EC_EARS,EC_OTHER,EC_TEETH,EC_NECK,EC_CHEST,EC_UPPER_BACK,"
SQLQ = SQLQ & "EC_LOWER_BACK,EC_ABDOMEN,EC_PELVIS,EC_LFT_SHOULDER,EC_LFT_ARM,EC_LFT_ELBOW,EC_LFT_FOREARM,"
SQLQ = SQLQ & "EC_RGT_SHOULDER,EC_RGT_ARM,EC_RGT_ELBOW,EC_RGT_FOREARM,EC_LFT_WRIST,EC_LFT_HAND,EC_LFT_FINGER,"
SQLQ = SQLQ & "EC_RGT_WRIST,EC_RGT_HAND,EC_RGT_FINGER,EC_LFT_HIP,EC_LFT_THIGH,EC_LFT_KNEE,EC_LFT_LOWER_LEG,"
SQLQ = SQLQ & "EC_RGT_HIP,EC_RGT_THIGH,EC_RGT_KNEE,EC_RGT_LOWER_LEG,EC_LFT_ANKLE,EC_LFT_FOOT,EC_LFT_TOES,"
SQLQ = SQLQ & "EC_RGT_ANKLE,EC_RGT_FOOT,EC_RGT_TOES,EC_PREMISES,EC_EMP_PREMISES,EC_OUTSIDE_PROV,EC_OUTSIDE_CITY,"
SQLQ = SQLQ & "EC_WITNESS,EC_WITNESS1_EMPNBR,EC_WITNESS1,EC_WITNESS2_EMPNBR,EC_WITNESS2,EC_INDIV_RESP,"
SQLQ = SQLQ & "EC_INDIV_NAME,EC_INDIV_PHONE,EC_SIMILAR_INJ,EC_SIMILAR_INJ_DEATAILS,EC_ANY_CONCERNS,EC_ANY_CONCERNS_DOC_KEY"

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
If glbLinamar Then 'Ticket #15172
    SQLQ = SQLQ & ",EC_OSHA300,EC_OSHACOM"
End If
FldList = SQLQ
End Function

Private Function FldList1()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "* "

If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList1 = SQLQ
End Function

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    Dim xSat As Integer
    Dim xSunDate, xSatDate

    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Me.cmdModify_Click
        Exit Sub
    End If
    
    
    If glbtermopen Then
        SQLQ = "SELECT " & FldList & " FROM Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & "WHERE EC_CASE=" & Data1.Recordset!EC_CASE
        If glbWFC Then
            'SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
            SQLQ = SQLQ & " AND TERM_SEQ=" & glbTERM_Seq
        End If
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT " & FldList & " FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & "WHERE EC_CASE = " & Data1.Recordset!EC_CASE
        If glbWFC Or glbCompSerial = "S/N - 2335W" Then
            SQLQ = SQLQ & " AND EC_EMPNBR =" & glbLEE_ID
        End If
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
       
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    

    'Form 7 - Additional Sections
    Call Additional_Form7Sections_Update
'    If Not Data1.Recordset.EOF Then
'        If rsDATA1.State <> 0 Then: If rsDATA1.EOF Then rsDATA1.Close Else If rsDATA1.EditMode = adEditAdd Then rsDATA1.CancelUpdate: rsDATA1.Close Else rsDATA1.Close
'        If glbtermopen Then
'            SQLQ = "SELECT " & FldList1
'            SQLQ = SQLQ & " FROM Term_OHS_FORM7_SECTIONS"
'            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'            SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
'            rsDATA1.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'            Data2.RecordSource = SQLQ
'        Else
'            SQLQ = "SELECT " & FldList1
'            SQLQ = SQLQ & " FROM HR_OHS_FORM7_SECTIONS"
'            SQLQ = SQLQ & " WHERE F7_EMPNBR = " & glbLEE_ID
'            SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
'            rsDATA1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'            Data2.RecordSource = SQLQ
'        End If
'        If rsDATA1.EOF Then
'            rsDATA1.AddNew
'            rsDATA1("F7_COMPNO") = "001"
'            rsDATA1("F7_EMPNBR") = glbLEE_ID
'            rsDATA1("F7_CASE") = Data1.Recordset!EC_CASE
'            rsDATA1("F7_FED_AMT") = GetEmpData(glbLEE_ID, "ED_TD1DOL", "")
'            rsDATA1("F7_PROV_AMT") = GetEmpData(glbLEE_ID, "ED_PROVAMT", "")
'            If glbtermopen Then
'                rsDATA1("TERM_SEQ") = glbTERM_Seq
'            End If
'
'            'Week 1 - 4 From/To Dates - for Section H-8 and I-C
'            'Week is Sun - Sat.
'            If IsDate(Data1.Recordset!EC_OCCDATE) Then
'
'                'Get which day of the week it is and get the # of days before the last Saturday
'                Select Case Weekday(Data1.Recordset!EC_OCCDATE)
'                    Case vbMonday
'                        xSat = -2
'                    Case vbTuesday
'                        xSat = -3
'                    Case vbWednesday
'                        xSat = -4
'                    Case vbThursday
'                        xSat = -5
'                    Case vbFriday
'                        xSat = -6
'                    Case vbSaturday
'                        xSat = -7
'                    Case vbSunday
'                        xSat = -8
'                End Select
'
'                'Week 1 - From
'                'Compute the date of the last Saturday
'                xSatDate = DateAdd("d", xSat, Data1.Recordset!EC_OCCDATE)
'
'                'Compute the date of last Sunday - 1 week prior
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK1") = xSunDate
'                rsDATA1("F7_FWEEK1") = xSunDate
'
'                'Week 1 - To
'                rsDATA1("F7_OTH_EARN_TO_WK1") = xSatDate
'                rsDATA1("F7_TWEEK1") = xSatDate
'
'
'                'Week 2 - From
'                xSatDate = DateAdd("d", -1, xSunDate)
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK2") = xSunDate
'                rsDATA1("F7_FWEEK2") = xSunDate
'
'                'Week 2 - To
'                rsDATA1("F7_OTH_EARN_TO_WK2") = xSatDate
'                rsDATA1("F7_TWEEK2") = xSatDate
'
'
'                'Week 3 - From
'                xSatDate = DateAdd("d", -1, xSunDate)
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK3") = xSunDate
'                rsDATA1("F7_FWEEK3") = xSunDate
'
'                'Week 3 - To
'                rsDATA1("F7_OTH_EARN_TO_WK3") = xSatDate
'                rsDATA1("F7_TWEEK3") = xSatDate
'
'
'                'Week 4 - From
'                xSatDate = DateAdd("d", -1, xSunDate)
'                xSunDate = DateAdd("d", -6, xSatDate)
'                rsDATA1("F7_OTH_EARN_FROM_WK4") = xSunDate
'                rsDATA1("F7_FWEEK4") = xSunDate
'
'                'Week 4 - To
'                rsDATA1("F7_OTH_EARN_TO_WK4") = xSatDate
'                rsDATA1("F7_TWEEK4") = xSatDate
'            End If
'
'            rsDATA1.Update
'        End If
'    End If
    
    Call SET_UP_MODE
    
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
UpdateRight = gSec_Upd_HSW7Injury And glbWSIBModule
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

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum

If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then 'Or rsDATA1.EOF Then
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

Private Sub lblEEID_Change()

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Injury WSIB Form 7 - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If

Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSINJURYWF7.Caption = "Injury WSIB Form 7 - " & Left$(glbLEE_SName, 5)
        frmEHSINJURYWF7.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
    
End If
End Sub

Private Function First_Body_Part_Selected()
    If chkHead.Value = 1 Then
        First_Body_Part_Selected = "HEAD"
    ElseIf chkFace.Value = 1 Then
        First_Body_Part_Selected = "FACE"
    ElseIf chkEyes.Value = 1 Then
        First_Body_Part_Selected = "EYE"
    ElseIf chkEars.Value = 1 Then
        First_Body_Part_Selected = "EAR"
    ElseIf chkTeeth.Value = 1 Then
        First_Body_Part_Selected = "TETH"
    ElseIf chkNeck.Value = 1 Then
        First_Body_Part_Selected = "NECK"
    ElseIf chkChest.Value = 1 Then
        First_Body_Part_Selected = "CHST"
    ElseIf chkUpBack.Value = 1 Then
        First_Body_Part_Selected = "UBCK"
    ElseIf chkLowBack.Value = 1 Then
        First_Body_Part_Selected = "LBCK"
    ElseIf chkAbdomen.Value = 1 Then
        First_Body_Part_Selected = "ABMN"
    ElseIf chkPelvis.Value = 1 Then
        First_Body_Part_Selected = "PLVS"
    ElseIf chkLShoulder.Value = 1 Then
        First_Body_Part_Selected = "SHDL"
    ElseIf chkRShoulder.Value = 1 Then
        First_Body_Part_Selected = "SHDR"
    ElseIf chkLArm.Value = 1 Then
        First_Body_Part_Selected = "ARML"
    ElseIf chkRArm.Value = 1 Then
        First_Body_Part_Selected = "ARMR"
    ElseIf chkLElbow.Value = 1 Then
        First_Body_Part_Selected = "ELBL"
    ElseIf chkRElbow.Value = 1 Then
        First_Body_Part_Selected = "ELBR"
    ElseIf chkLForearm.Value = 1 Then
        First_Body_Part_Selected = "FAML"
    ElseIf chkRForearm.Value = 1 Then
        First_Body_Part_Selected = "FAMR"
    ElseIf chkLWrist.Value = 1 Then
        First_Body_Part_Selected = "WRTL"
    ElseIf chkRWrist.Value = 1 Then
        First_Body_Part_Selected = "WRTR"
    ElseIf chkLHand.Value = 1 Then
        First_Body_Part_Selected = "HNDL"
    ElseIf chkRHand.Value = 1 Then
        First_Body_Part_Selected = "HNDR"
    ElseIf chkLFingers.Value = 1 Then
        First_Body_Part_Selected = "FNGL"
    ElseIf chkRFingers.Value = 1 Then
        First_Body_Part_Selected = "FNGR"
    ElseIf chkLHip.Value = 1 Then
        First_Body_Part_Selected = "HIPL"
    ElseIf chkRHip.Value = 1 Then
        First_Body_Part_Selected = "HIPR"
    ElseIf chkLThigh.Value = 1 Then
        First_Body_Part_Selected = "THGL"
    ElseIf chkRThigh.Value = 1 Then
        First_Body_Part_Selected = "THGR"
    ElseIf chkLKnee.Value = 1 Then
        First_Body_Part_Selected = "KNEL"
    ElseIf chkRKnee.Value = 1 Then
        First_Body_Part_Selected = "KNER"
    ElseIf chkLLowLeg.Value = 1 Then
        First_Body_Part_Selected = "LLGL"
    ElseIf chkRLowLeg.Value = 1 Then
        First_Body_Part_Selected = "LLGR"
    ElseIf chkLAnkle.Value = 1 Then
        First_Body_Part_Selected = "ANKL"
    ElseIf chkRAnkle.Value = 1 Then
        First_Body_Part_Selected = "ANKR"
    ElseIf chkLFoot.Value = 1 Then
        First_Body_Part_Selected = "FOTL"
    ElseIf chkRFoot.Value = 1 Then
        First_Body_Part_Selected = "FOTR"
    ElseIf chkLToes.Value = 1 Then
        First_Body_Part_Selected = "TOEL"
    ElseIf chkRToes.Value = 1 Then
        First_Body_Part_Selected = "TOER"
    ElseIf chkOther.Value = 1 Then
        First_Body_Part_Selected = clpCode(0).Text
    End If
End Function

Private Sub Additional_Form7Sections_Update()
    Dim SQLQ As String
    Dim xSat As Integer
    Dim xSunDate As Date
    Dim xSatDate As Date

    'Form 7 - Additional Sections
    If Not Data1.Recordset.EOF Then
        If rsDATA1.State <> 0 Then: If rsDATA1.EOF Then rsDATA1.Close Else If rsDATA1.EditMode = adEditAdd Then rsDATA1.CancelUpdate: rsDATA1.Close Else rsDATA1.Close
        If glbtermopen Then
            SQLQ = "SELECT " & FldList1
            SQLQ = SQLQ & " FROM Term_OHS_FORM7_SECTIONS"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
            SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
            rsDATA1.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            Data2.RecordSource = SQLQ
        Else
            SQLQ = "SELECT " & FldList1
            SQLQ = SQLQ & " FROM HR_OHS_FORM7_SECTIONS"
            SQLQ = SQLQ & " WHERE F7_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND F7_CASE = " & Data1.Recordset!EC_CASE
            rsDATA1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Data2.RecordSource = SQLQ
        End If
        If rsDATA1.EOF Then
            rsDATA1.AddNew
            rsDATA1("F7_COMPNO") = "001"
            rsDATA1("F7_EMPNBR") = glbLEE_ID
            rsDATA1("F7_CASE") = Data1.Recordset!EC_CASE
            rsDATA1("F7_FED_AMT") = GetEmpData(glbLEE_ID, "ED_TD1DOL", "")
            rsDATA1("F7_PROV_AMT") = GetEmpData(glbLEE_ID, "ED_PROVAMT", "")
            If glbtermopen Then
                rsDATA1("TERM_SEQ") = glbTERM_Seq
            End If

            'Week 1 - 4 From/To Dates - for Section H-8 and I-C
            'Week is Sun - Sat.
            If IsDate(Data1.Recordset!EC_OCCDATE) Then

                'Get which day of the week it is and get the # of days before the last Saturday
                Select Case Weekday(Data1.Recordset!EC_OCCDATE)
                    Case vbMonday
                        xSat = -2
                    Case vbTuesday
                        xSat = -3
                    Case vbWednesday
                        xSat = -4
                    Case vbThursday
                        xSat = -5
                    Case vbFriday
                        xSat = -6
                    Case vbSaturday
                        xSat = -7
                    Case vbSunday
                        xSat = -8
                End Select

                'Week 1 - From
                'Compute the date of the last Saturday
                'xSatDate = DateAdd("d", xSat, Data1.Recordset!EC_OCCDATE)
                'Date prior to Incident Date
                xSatDate = DateAdd("d", -1, CVDate(Data1.Recordset!EC_OCCDATE))

                'Compute the date of last Sunday - 1 week prior
                xSunDate = DateAdd("d", -6, xSatDate)
                rsDATA1("F7_OTH_EARN_FROM_WK1") = xSunDate
                rsDATA1("F7_FWEEK1") = xSunDate

                'Week 1 - To
                rsDATA1("F7_OTH_EARN_TO_WK1") = xSatDate
                rsDATA1("F7_TWEEK1") = xSatDate


                'Week 2 - From
                xSatDate = DateAdd("d", -1, xSunDate)
                xSunDate = DateAdd("d", -6, xSatDate)
                rsDATA1("F7_OTH_EARN_FROM_WK2") = xSunDate
                rsDATA1("F7_FWEEK2") = xSunDate

                'Week 2 - To
                rsDATA1("F7_OTH_EARN_TO_WK2") = xSatDate
                rsDATA1("F7_TWEEK2") = xSatDate


                'Week 3 - From
                xSatDate = DateAdd("d", -1, xSunDate)
                xSunDate = DateAdd("d", -6, xSatDate)
                rsDATA1("F7_OTH_EARN_FROM_WK3") = xSunDate
                rsDATA1("F7_FWEEK3") = xSunDate

                'Week 3 - To
                rsDATA1("F7_OTH_EARN_TO_WK3") = xSatDate
                rsDATA1("F7_TWEEK3") = xSatDate


                'Week 4 - From
                xSatDate = DateAdd("d", -1, xSunDate)
                xSunDate = DateAdd("d", -6, xSatDate)
                rsDATA1("F7_OTH_EARN_FROM_WK4") = xSunDate
                rsDATA1("F7_FWEEK4") = xSunDate

                'Week 4 - To
                rsDATA1("F7_OTH_EARN_TO_WK4") = xSatDate
                rsDATA1("F7_TWEEK4") = xSatDate
            End If

            rsDATA1.Update
        End If
    End If


End Sub

