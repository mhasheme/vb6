VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPayPeriodMaster 
   Caption         =   "Pay Period Master"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmPayPeriodMaster.frx":0000
      Height          =   2085
      Left            =   180
      OleObjectBlob   =   "frmPayPeriodMaster.frx":0014
      TabIndex        =   0
      Top             =   60
      Width           =   9765
   End
   Begin VB.VScrollBar scrControl 
      Height          =   3975
      LargeChange     =   1000
      Left            =   9600
      SmallChange     =   100
      TabIndex        =   252
      Top             =   3810
      Width           =   315
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "frmTop"
      Height          =   1695
      Left            =   240
      TabIndex        =   253
      Top             =   2120
      Width           =   9645
      Begin VB.ComboBox comNbrOfPayPeriod 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtYear 
         Appearance      =   0  'Flat
         DataField       =   "PP_YEAR"
         Height          =   285
         Left            =   1755
         TabIndex        =   1
         Tag             =   "61-Year"
         Top             =   90
         Width           =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpPAYP 
         DataField       =   "PP_PAYP"
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Tag             =   "00-Enter Pay Period Code"
         Top             =   390
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDPP"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   6060
         TabIndex        =   4
         Tag             =   "00-Enter Union Code"
         Top             =   390
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Tag             =   "00-Specific Department Desired"
         Top             =   990
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Tag             =   "00-Specific Division Desired"
         Top             =   690
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   6060
         TabIndex        =   8
         Tag             =   "00-Section - Code"
         Top             =   990
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   6060
         TabIndex        =   6
         Tag             =   "00-Enter Location Code"
         Top             =   690
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin VB.Label lblNbrOfPayPeriod 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Pay Period"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4350
         TabIndex        =   265
         Top             =   120
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lblStart 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date:"
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
         Height          =   255
         Left            =   1500
         TabIndex        =   264
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date:"
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
         Height          =   255
         Left            =   4590
         TabIndex        =   263
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lbUpload 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
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
         Height          =   255
         Left            =   7080
         TabIndex        =   262
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period"
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
         Height          =   255
         Left            =   30
         TabIndex        =   261
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
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
         Height          =   255
         Left            =   90
         TabIndex        =   260
         Top             =   120
         Width           =   645
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Period Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   259
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   258
         Top             =   720
         Width           =   555
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   257
         Top             =   1020
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5475
         TabIndex        =   256
         Top             =   450
         Width           =   420
      End
      Begin VB.Label lblSection 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5355
         TabIndex        =   255
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label lblLocation 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5250
         TabIndex        =   254
         Top             =   750
         Width           =   675
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   250
      Top             =   7770
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
      _ExtentY        =   979
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
      Begin VB.CommandButton cmdPrintAll 
         Appearance      =   0  'Flat
         Caption         =   "P&rint ALL"
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
         Left            =   570
         TabIndex        =   251
         Tag             =   "Print Listing "
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin MSAdodcLib.Adodc datPP 
         Height          =   330
         Left            =   8100
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
         Caption         =   "datPP"
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
         Left            =   7200
         Top             =   60
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   0
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
   End
   Begin VB.Frame VacFram 
      BorderStyle     =   0  'None
      Height          =   16000
      Left            =   240
      TabIndex        =   249
      Top             =   3810
      Width           =   9105
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   7200
         TabIndex        =   140
         Tag             =   "40-Uploaded -y/n"
         Top             =   9600
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   7200
         TabIndex        =   144
         Tag             =   "40-Uploaded -y/n"
         Top             =   9900
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   7200
         TabIndex        =   148
         Tag             =   "40-Uploaded -y/n"
         Top             =   10200
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   7200
         TabIndex        =   152
         Tag             =   "40-Uploaded -y/n"
         Top             =   10500
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   7200
         TabIndex        =   156
         Tag             =   "40-Uploaded -y/n"
         Top             =   10800
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   7200
         TabIndex        =   160
         Tag             =   "40-Uploaded -y/n"
         Top             =   11100
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   7200
         TabIndex        =   164
         Tag             =   "40-Uploaded -y/n"
         Top             =   11400
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   7200
         TabIndex        =   168
         Tag             =   "40-Uploaded -y/n"
         Top             =   11700
         Width           =   315
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   32
         Left            =   90
         TabIndex        =   137
         Top             =   9600
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   33
         Left            =   90
         TabIndex        =   141
         Top             =   9900
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   34
         Left            =   90
         TabIndex        =   145
         Top             =   10200
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   35
         Left            =   90
         TabIndex        =   149
         Top             =   10500
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   36
         Left            =   90
         TabIndex        =   153
         Top             =   10800
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   37
         Left            =   90
         TabIndex        =   157
         Top             =   11100
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   38
         Left            =   90
         TabIndex        =   161
         Top             =   11400
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   39
         Left            =   90
         TabIndex        =   165
         Top             =   11700
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   47
         Left            =   90
         TabIndex        =   197
         Top             =   14100
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   46
         Left            =   90
         TabIndex        =   193
         Top             =   13800
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   45
         Left            =   90
         TabIndex        =   189
         Top             =   13500
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   44
         Left            =   90
         TabIndex        =   185
         Top             =   13200
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   43
         Left            =   90
         TabIndex        =   181
         Top             =   12900
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   42
         Left            =   90
         TabIndex        =   177
         Top             =   12600
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   41
         Left            =   90
         TabIndex        =   173
         Top             =   12300
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   40
         Left            =   90
         TabIndex        =   169
         Top             =   12000
         Width           =   555
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   7200
         TabIndex        =   200
         Tag             =   "40-Uploaded -y/n"
         Top             =   14100
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   7200
         TabIndex        =   196
         Tag             =   "40-Uploaded -y/n"
         Top             =   13800
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   7200
         TabIndex        =   192
         Tag             =   "40-Uploaded -y/n"
         Top             =   13500
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   7200
         TabIndex        =   188
         Tag             =   "40-Uploaded -y/n"
         Top             =   13200
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   7200
         TabIndex        =   184
         Tag             =   "40-Uploaded -y/n"
         Top             =   12900
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   7200
         TabIndex        =   180
         Tag             =   "40-Uploaded -y/n"
         Top             =   12600
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   7200
         TabIndex        =   176
         Tag             =   "40-Uploaded -y/n"
         Top             =   12300
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   7200
         TabIndex        =   172
         Tag             =   "40-Uploaded -y/n"
         Top             =   12000
         Width           =   315
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   55
         Left            =   90
         TabIndex        =   229
         Top             =   16500
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   54
         Left            =   90
         TabIndex        =   225
         Top             =   16200
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   53
         Left            =   90
         TabIndex        =   221
         Top             =   15900
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   52
         Left            =   90
         TabIndex        =   217
         Top             =   15600
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   51
         Left            =   90
         TabIndex        =   213
         Top             =   15300
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   50
         Left            =   90
         TabIndex        =   209
         Top             =   15000
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   49
         Left            =   90
         TabIndex        =   205
         Top             =   14700
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   48
         Left            =   90
         TabIndex        =   201
         Top             =   14400
         Width           =   555
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   7200
         TabIndex        =   232
         Tag             =   "40-Uploaded -y/n"
         Top             =   16500
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   7200
         TabIndex        =   228
         Tag             =   "40-Uploaded -y/n"
         Top             =   16200
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   7200
         TabIndex        =   224
         Tag             =   "40-Uploaded -y/n"
         Top             =   15900
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   7200
         TabIndex        =   220
         Tag             =   "40-Uploaded -y/n"
         Top             =   15600
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   7200
         TabIndex        =   216
         Tag             =   "40-Uploaded -y/n"
         Top             =   15300
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   7200
         TabIndex        =   212
         Tag             =   "40-Uploaded -y/n"
         Top             =   15000
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   7200
         TabIndex        =   208
         Tag             =   "40-Uploaded -y/n"
         Top             =   14700
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   7200
         TabIndex        =   204
         Tag             =   "40-Uploaded -y/n"
         Top             =   14400
         Width           =   315
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   59
         Left            =   90
         TabIndex        =   245
         Top             =   17700
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   58
         Left            =   90
         TabIndex        =   241
         Top             =   17400
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   57
         Left            =   90
         TabIndex        =   237
         Top             =   17100
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   56
         Left            =   90
         TabIndex        =   233
         Top             =   16800
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   7200
         TabIndex        =   248
         Tag             =   "40-Uploaded -y/n"
         Top             =   17700
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   7200
         TabIndex        =   244
         Tag             =   "40-Uploaded -y/n"
         Top             =   17400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   7200
         TabIndex        =   240
         Tag             =   "40-Uploaded -y/n"
         Top             =   17100
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   7200
         TabIndex        =   236
         Tag             =   "40-Uploaded -y/n"
         Top             =   16800
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   7200
         TabIndex        =   108
         Tag             =   "40-Uploaded -y/n"
         Top             =   7200
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   7200
         TabIndex        =   112
         Tag             =   "40-Uploaded -y/n"
         Top             =   7500
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   7200
         TabIndex        =   116
         Tag             =   "40-Uploaded -y/n"
         Top             =   7800
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   7200
         TabIndex        =   120
         Tag             =   "40-Uploaded -y/n"
         Top             =   8100
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   7200
         TabIndex        =   124
         Tag             =   "40-Uploaded -y/n"
         Top             =   8400
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   7200
         TabIndex        =   128
         Tag             =   "40-Uploaded -y/n"
         Top             =   8700
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   7200
         TabIndex        =   132
         Tag             =   "40-Uploaded -y/n"
         Top             =   9000
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   7200
         TabIndex        =   136
         Tag             =   "40-Uploaded -y/n"
         Top             =   9300
         Width           =   315
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   24
         Left            =   90
         TabIndex        =   105
         Top             =   7200
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   25
         Left            =   90
         TabIndex        =   109
         Top             =   7500
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   26
         Left            =   90
         TabIndex        =   113
         Top             =   7800
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   27
         Left            =   90
         TabIndex        =   117
         Top             =   8100
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   28
         Left            =   90
         TabIndex        =   121
         Top             =   8400
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   29
         Left            =   90
         TabIndex        =   125
         Top             =   8700
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   30
         Left            =   90
         TabIndex        =   129
         Top             =   9000
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   31
         Left            =   90
         TabIndex        =   133
         Top             =   9300
         Width           =   555
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   7200
         TabIndex        =   76
         Tag             =   "40-Uploaded -y/n"
         Top             =   4800
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   7200
         TabIndex        =   80
         Tag             =   "40-Uploaded -y/n"
         Top             =   5100
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   7200
         TabIndex        =   84
         Tag             =   "40-Uploaded -y/n"
         Top             =   5400
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   7200
         TabIndex        =   88
         Tag             =   "40-Uploaded -y/n"
         Top             =   5700
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   7200
         TabIndex        =   92
         Tag             =   "40-Uploaded -y/n"
         Top             =   6000
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   7200
         TabIndex        =   96
         Tag             =   "40-Uploaded -y/n"
         Top             =   6300
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   7200
         TabIndex        =   100
         Tag             =   "40-Uploaded -y/n"
         Top             =   6600
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   7200
         TabIndex        =   104
         Tag             =   "40-Uploaded -y/n"
         Top             =   6900
         Width           =   315
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   90
         TabIndex        =   73
         Top             =   4800
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   90
         TabIndex        =   77
         Top             =   5100
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   18
         Left            =   90
         TabIndex        =   81
         Top             =   5400
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   19
         Left            =   90
         TabIndex        =   85
         Top             =   5700
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   20
         Left            =   90
         TabIndex        =   89
         Top             =   6000
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   21
         Left            =   90
         TabIndex        =   93
         Top             =   6300
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   22
         Left            =   90
         TabIndex        =   97
         Top             =   6600
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   23
         Left            =   90
         TabIndex        =   101
         Top             =   6900
         Width           =   555
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7200
         TabIndex        =   44
         Tag             =   "40-Uploaded -y/n"
         Top             =   2400
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   7200
         TabIndex        =   48
         Tag             =   "40-Uploaded -y/n"
         Top             =   2700
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   7200
         TabIndex        =   52
         Tag             =   "40-Uploaded -y/n"
         Top             =   3000
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   7200
         TabIndex        =   56
         Tag             =   "40-Uploaded -y/n"
         Top             =   3300
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   7200
         TabIndex        =   60
         Tag             =   "40-Uploaded -y/n"
         Top             =   3600
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   7200
         TabIndex        =   64
         Tag             =   "40-Uploaded -y/n"
         Top             =   3900
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   7200
         TabIndex        =   68
         Tag             =   "40-Uploaded -y/n"
         Top             =   4200
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   7200
         TabIndex        =   72
         Tag             =   "40-Uploaded -y/n"
         Top             =   4500
         Width           =   315
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   41
         Top             =   2400
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   45
         Top             =   2700
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   49
         Top             =   3000
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   53
         Top             =   3300
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   90
         TabIndex        =   57
         Top             =   3600
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   13
         Left            =   90
         TabIndex        =   61
         Top             =   3900
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   14
         Left            =   90
         TabIndex        =   65
         Top             =   4200
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   15
         Left            =   90
         TabIndex        =   69
         Top             =   4500
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   37
         Top             =   2100
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   33
         Top             =   1800
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   29
         Top             =   1500
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   900
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   600
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   555
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   0
         Width           =   555
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7200
         TabIndex        =   40
         Tag             =   "40-Uploaded -y/n"
         Top             =   2100
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   7200
         TabIndex        =   36
         Tag             =   "40-Uploaded -y/n"
         Top             =   1800
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7200
         TabIndex        =   32
         Tag             =   "40-Uploaded -y/n"
         Top             =   1500
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7200
         TabIndex        =   28
         Tag             =   "40-Uploaded -y/n"
         Top             =   1200
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   24
         Tag             =   "40-Uploaded -y/n"
         Top             =   900
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   20
         Tag             =   "40-Uploaded -y/n"
         Top             =   600
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   16
         Tag             =   "40-Uploaded -y/n"
         Top             =   300
         Width           =   315
      End
      Begin VB.CheckBox chkUploaded 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   12
         Tag             =   "40-Uploaded -y/n"
         Top             =   0
         Width           =   315
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   0
         Left            =   4260
         TabIndex        =   11
         Tag             =   "41-End Date"
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   10
         Tag             =   "41-Start Date"
         Top             =   0
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   1
         Left            =   4260
         TabIndex        =   15
         Tag             =   "41-End Date"
         Top             =   300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   14
         Tag             =   "41-Start Date"
         Top             =   300
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   2
         Left            =   4260
         TabIndex        =   19
         Tag             =   "41-End Date"
         Top             =   600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   2
         Left            =   1140
         TabIndex        =   18
         Tag             =   "41-Start Date"
         Top             =   600
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   3
         Left            =   4260
         TabIndex        =   23
         Tag             =   "41-End Date"
         Top             =   900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   3
         Left            =   1140
         TabIndex        =   22
         Tag             =   "41-Start Date"
         Top             =   900
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   4
         Left            =   4260
         TabIndex        =   27
         Tag             =   "41-End Date"
         Top             =   1200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   4
         Left            =   1140
         TabIndex        =   26
         Tag             =   "41-Start Date"
         Top             =   1200
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   5
         Left            =   4260
         TabIndex        =   31
         Tag             =   "41-End Date"
         Top             =   1500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   5
         Left            =   1140
         TabIndex        =   30
         Tag             =   "41-Start Date"
         Top             =   1500
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   6
         Left            =   4260
         TabIndex        =   35
         Tag             =   "41-End Date"
         Top             =   1800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   6
         Left            =   1140
         TabIndex        =   34
         Tag             =   "41-Start Date"
         Top             =   1800
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   7
         Left            =   4260
         TabIndex        =   39
         Tag             =   "41-End Date"
         Top             =   2100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   7
         Left            =   1140
         TabIndex        =   38
         Tag             =   "41-Start Date"
         Top             =   2100
         Width           =   2500
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   8
         Left            =   4260
         TabIndex        =   43
         Tag             =   "41-End Date"
         Top             =   2400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   8
         Left            =   1140
         TabIndex        =   42
         Tag             =   "41-Start Date"
         Top             =   2400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   9
         Left            =   4260
         TabIndex        =   47
         Tag             =   "41-End Date"
         Top             =   2700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   9
         Left            =   1140
         TabIndex        =   46
         Tag             =   "41-Start Date"
         Top             =   2700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   10
         Left            =   4260
         TabIndex        =   51
         Tag             =   "41-End Date"
         Top             =   3000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   10
         Left            =   1140
         TabIndex        =   50
         Tag             =   "41-Start Date"
         Top             =   3000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   11
         Left            =   4260
         TabIndex        =   55
         Tag             =   "41-End Date"
         Top             =   3300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   11
         Left            =   1140
         TabIndex        =   54
         Tag             =   "41-Start Date"
         Top             =   3300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   12
         Left            =   4260
         TabIndex        =   59
         Tag             =   "41-End Date"
         Top             =   3600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   12
         Left            =   1140
         TabIndex        =   58
         Tag             =   "41-Start Date"
         Top             =   3600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   13
         Left            =   4260
         TabIndex        =   63
         Tag             =   "41-End Date"
         Top             =   3900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   13
         Left            =   1140
         TabIndex        =   62
         Tag             =   "41-Start Date"
         Top             =   3900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   14
         Left            =   4260
         TabIndex        =   67
         Tag             =   "41-End Date"
         Top             =   4200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   14
         Left            =   1140
         TabIndex        =   66
         Tag             =   "41-Start Date"
         Top             =   4200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   15
         Left            =   4260
         TabIndex        =   71
         Tag             =   "41-End Date"
         Top             =   4500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   15
         Left            =   1140
         TabIndex        =   70
         Tag             =   "41-Start Date"
         Top             =   4500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   16
         Left            =   4260
         TabIndex        =   75
         Tag             =   "41-End Date"
         Top             =   4800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   16
         Left            =   1140
         TabIndex        =   74
         Tag             =   "41-Start Date"
         Top             =   4800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   17
         Left            =   4260
         TabIndex        =   79
         Tag             =   "41-End Date"
         Top             =   5100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   17
         Left            =   1140
         TabIndex        =   78
         Tag             =   "41-Start Date"
         Top             =   5100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   18
         Left            =   4260
         TabIndex        =   83
         Tag             =   "41-End Date"
         Top             =   5400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   18
         Left            =   1140
         TabIndex        =   82
         Tag             =   "41-Start Date"
         Top             =   5400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   19
         Left            =   4260
         TabIndex        =   87
         Tag             =   "41-End Date"
         Top             =   5700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   19
         Left            =   1140
         TabIndex        =   86
         Tag             =   "41-Start Date"
         Top             =   5700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   20
         Left            =   4260
         TabIndex        =   91
         Tag             =   "41-End Date"
         Top             =   6000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   20
         Left            =   1140
         TabIndex        =   90
         Tag             =   "41-Start Date"
         Top             =   6000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   21
         Left            =   4260
         TabIndex        =   95
         Tag             =   "41-End Date"
         Top             =   6300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   21
         Left            =   1140
         TabIndex        =   94
         Tag             =   "41-Start Date"
         Top             =   6300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   22
         Left            =   4260
         TabIndex        =   99
         Tag             =   "41-End Date"
         Top             =   6600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   22
         Left            =   1140
         TabIndex        =   98
         Tag             =   "41-Start Date"
         Top             =   6600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   23
         Left            =   4260
         TabIndex        =   103
         Tag             =   "41-End Date"
         Top             =   6900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   23
         Left            =   1140
         TabIndex        =   102
         Tag             =   "41-Start Date"
         Top             =   6900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   24
         Left            =   4260
         TabIndex        =   107
         Tag             =   "41-End Date"
         Top             =   7200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   24
         Left            =   1140
         TabIndex        =   106
         Tag             =   "41-Start Date"
         Top             =   7200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   25
         Left            =   4260
         TabIndex        =   111
         Tag             =   "41-End Date"
         Top             =   7500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   25
         Left            =   1140
         TabIndex        =   110
         Tag             =   "41-Start Date"
         Top             =   7500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   26
         Left            =   4260
         TabIndex        =   115
         Tag             =   "41-End Date"
         Top             =   7800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   26
         Left            =   1140
         TabIndex        =   114
         Tag             =   "41-Start Date"
         Top             =   7800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   27
         Left            =   4260
         TabIndex        =   119
         Tag             =   "41-End Date"
         Top             =   8100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   27
         Left            =   1140
         TabIndex        =   118
         Tag             =   "41-Start Date"
         Top             =   8100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   28
         Left            =   4260
         TabIndex        =   123
         Tag             =   "41-End Date"
         Top             =   8400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   28
         Left            =   1140
         TabIndex        =   122
         Tag             =   "41-Start Date"
         Top             =   8400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   29
         Left            =   4260
         TabIndex        =   127
         Tag             =   "41-End Date"
         Top             =   8700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   29
         Left            =   1140
         TabIndex        =   126
         Tag             =   "41-Start Date"
         Top             =   8700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   30
         Left            =   4260
         TabIndex        =   131
         Tag             =   "41-End Date"
         Top             =   9000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   30
         Left            =   1140
         TabIndex        =   130
         Tag             =   "41-Start Date"
         Top             =   9000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   31
         Left            =   4260
         TabIndex        =   135
         Tag             =   "41-End Date"
         Top             =   9300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   31
         Left            =   1140
         TabIndex        =   134
         Tag             =   "41-Start Date"
         Top             =   9300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   32
         Left            =   4260
         TabIndex        =   139
         Tag             =   "41-End Date"
         Top             =   9600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   32
         Left            =   1140
         TabIndex        =   138
         Tag             =   "41-Start Date"
         Top             =   9600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   33
         Left            =   4260
         TabIndex        =   143
         Tag             =   "41-End Date"
         Top             =   9900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   33
         Left            =   1140
         TabIndex        =   142
         Tag             =   "41-Start Date"
         Top             =   9900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   34
         Left            =   4260
         TabIndex        =   147
         Tag             =   "41-End Date"
         Top             =   10200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   34
         Left            =   1140
         TabIndex        =   146
         Tag             =   "41-Start Date"
         Top             =   10200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   35
         Left            =   4260
         TabIndex        =   151
         Tag             =   "41-End Date"
         Top             =   10500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   35
         Left            =   1140
         TabIndex        =   150
         Tag             =   "41-Start Date"
         Top             =   10500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   36
         Left            =   4260
         TabIndex        =   155
         Tag             =   "41-End Date"
         Top             =   10800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   36
         Left            =   1140
         TabIndex        =   154
         Tag             =   "41-Start Date"
         Top             =   10800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   37
         Left            =   4260
         TabIndex        =   159
         Tag             =   "41-End Date"
         Top             =   11100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   37
         Left            =   1140
         TabIndex        =   158
         Tag             =   "41-Start Date"
         Top             =   11100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   38
         Left            =   4260
         TabIndex        =   163
         Tag             =   "41-End Date"
         Top             =   11400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   38
         Left            =   1140
         TabIndex        =   162
         Tag             =   "41-Start Date"
         Top             =   11400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   39
         Left            =   4260
         TabIndex        =   167
         Tag             =   "41-End Date"
         Top             =   11700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   39
         Left            =   1140
         TabIndex        =   166
         Tag             =   "41-Start Date"
         Top             =   11700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   40
         Left            =   4260
         TabIndex        =   171
         Tag             =   "41-End Date"
         Top             =   12000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   40
         Left            =   1140
         TabIndex        =   170
         Tag             =   "41-Start Date"
         Top             =   12000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   41
         Left            =   4260
         TabIndex        =   175
         Tag             =   "41-End Date"
         Top             =   12300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   41
         Left            =   1140
         TabIndex        =   174
         Tag             =   "41-Start Date"
         Top             =   12300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   42
         Left            =   4260
         TabIndex        =   179
         Tag             =   "41-End Date"
         Top             =   12600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   42
         Left            =   1140
         TabIndex        =   178
         Tag             =   "41-Start Date"
         Top             =   12600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   43
         Left            =   4260
         TabIndex        =   183
         Tag             =   "41-End Date"
         Top             =   12900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   43
         Left            =   1140
         TabIndex        =   182
         Tag             =   "41-Start Date"
         Top             =   12900
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   44
         Left            =   4260
         TabIndex        =   187
         Tag             =   "41-End Date"
         Top             =   13200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   44
         Left            =   1140
         TabIndex        =   186
         Tag             =   "41-Start Date"
         Top             =   13200
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   45
         Left            =   4260
         TabIndex        =   191
         Tag             =   "41-End Date"
         Top             =   13500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   45
         Left            =   1140
         TabIndex        =   190
         Tag             =   "41-Start Date"
         Top             =   13500
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   46
         Left            =   4260
         TabIndex        =   195
         Tag             =   "41-End Date"
         Top             =   13800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   46
         Left            =   1140
         TabIndex        =   194
         Tag             =   "41-Start Date"
         Top             =   13800
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   47
         Left            =   4260
         TabIndex        =   199
         Tag             =   "41-End Date"
         Top             =   14100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   47
         Left            =   1140
         TabIndex        =   198
         Tag             =   "41-Start Date"
         Top             =   14100
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   48
         Left            =   4260
         TabIndex        =   203
         Tag             =   "41-End Date"
         Top             =   14400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   48
         Left            =   1140
         TabIndex        =   202
         Tag             =   "41-Start Date"
         Top             =   14400
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   49
         Left            =   4260
         TabIndex        =   207
         Tag             =   "41-End Date"
         Top             =   14700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   49
         Left            =   1140
         TabIndex        =   206
         Tag             =   "41-Start Date"
         Top             =   14700
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   50
         Left            =   4260
         TabIndex        =   211
         Tag             =   "41-End Date"
         Top             =   15000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   50
         Left            =   1140
         TabIndex        =   210
         Tag             =   "41-Start Date"
         Top             =   15000
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   51
         Left            =   4260
         TabIndex        =   215
         Tag             =   "41-End Date"
         Top             =   15300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   51
         Left            =   1140
         TabIndex        =   214
         Tag             =   "41-Start Date"
         Top             =   15300
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   52
         Left            =   4260
         TabIndex        =   219
         Tag             =   "41-End Date"
         Top             =   15600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   52
         Left            =   1140
         TabIndex        =   218
         Tag             =   "41-Start Date"
         Top             =   15600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   53
         Left            =   4260
         TabIndex        =   223
         Tag             =   "41-End Date"
         Top             =   15900
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   53
         Left            =   1140
         TabIndex        =   222
         Tag             =   "41-Start Date"
         Top             =   15900
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   54
         Left            =   4260
         TabIndex        =   227
         Tag             =   "41-End Date"
         Top             =   16200
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   54
         Left            =   1140
         TabIndex        =   226
         Tag             =   "41-Start Date"
         Top             =   16200
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   55
         Left            =   4260
         TabIndex        =   231
         Tag             =   "41-End Date"
         Top             =   16500
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   55
         Left            =   1140
         TabIndex        =   230
         Tag             =   "41-Start Date"
         Top             =   16500
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   56
         Left            =   4260
         TabIndex        =   235
         Tag             =   "41-End Date"
         Top             =   16800
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   56
         Left            =   1140
         TabIndex        =   234
         Tag             =   "41-Start Date"
         Top             =   16800
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   57
         Left            =   4260
         TabIndex        =   239
         Tag             =   "41-End Date"
         Top             =   17100
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   57
         Left            =   1140
         TabIndex        =   238
         Tag             =   "41-Start Date"
         Top             =   17100
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   58
         Left            =   4260
         TabIndex        =   243
         Tag             =   "41-End Date"
         Top             =   17400
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   58
         Left            =   1140
         TabIndex        =   242
         Tag             =   "41-Start Date"
         Top             =   17400
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Index           =   59
         Left            =   4260
         TabIndex        =   247
         Tag             =   "41-End Date"
         Top             =   17700
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Index           =   59
         Left            =   1140
         TabIndex        =   246
         Tag             =   "41-Start Date"
         Top             =   17700
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
   End
End
Attribute VB_Name = "frmPayPeriodMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fTablHREMP As New ADODB.Recordset         ' table view of HREMP
Dim oldStart
Dim Actn

Dim fglbCode$       ' are we dealing with Sick records?"
Dim fglbMaxRanges%
Dim glbFrmCaption$, glbErrNum&

Dim ControlsShown As Boolean
Dim ODIV, ODept, oOrg, oPayP
Dim OSection, OLoc, oYear

Dim FlagRefresh As Boolean

Dim fglbESQLQ, fglbVSQLQ
Dim fglbNew As Boolean
Dim fglbRunTimes
Dim Memplist1, Memplist2
Dim xPPNoChanged As String

Private Function chkPayPeriod()
Dim X%, Y%

chkPayPeriod = False

On Error GoTo chkPayPeriod_Err

If Len(txtYear) = 0 Then
    MsgBox "Year is required field"
    txtYear.SetFocus
    Exit Function
End If

For X% = 0 To 4
    If X <> 1 And X <> 2 Then
        If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
            MsgBox "If Code entered it must be known"
            clpCode(X%).SetFocus
            Exit Function
        End If
    End If
Next X%

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
If clpPAYP.Caption = "Unassigned" Then
    MsgBox "If " & clpPAYP.Caption & " Entered - it must be known"
    clpPAYP.SetFocus
    Exit Function
End If


If Len(txtSeq(0)) < 1 Then
    MsgBox "You must have at least one Pay Period."
    If txtSeq(0).Enabled Then txtSeq(0).SetFocus
    Exit Function
End If


fglbMaxRanges% = 0  ' 0 is first range

Dim intRangesSet%
intRangesSet% = 0    ' 1 to 4 with 0 implying none


For X% = 0 To txtSeq.count - 1
    If Len(txtSeq(X%)) > 0 Then
        If Not IsNumeric(txtSeq(X%)) Then
            MsgBox "Data Entered Must Be Numeric"
            txtSeq(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(dlpFrom(X%)) > 0 Then
        If Not IsDate(dlpFrom(X%)) Then
            MsgBox "Data Entered Must Be Date"
            dlpFrom(X%).SetFocus
            Exit Function
        End If
    End If
    If Len(dlpTo(X%)) > 0 Then
        If Not IsDate(dlpTo(X%)) Then
            MsgBox "Data Entered Must Be Date"
            dlpTo(X%).SetFocus
            Exit Function
        End If
    End If

    If X% > 0 And Len(txtSeq(X%)) > 0 Then
        If Val(txtSeq(X%)) < Val(txtSeq(X% - 1)) And Val(txtSeq(X%)) <> 0 Then
            MsgBox "Pay Period must be sequential"
            txtSeq(X%).SetFocus
            Exit Function
        End If
        If Len(dlpFrom(X)) = 0 Then
            MsgBox "Start Date must be entered"
            dlpFrom(X).SetFocus
            Exit Function
        End If
        If Len(dlpTo(X)) = 0 Then
            MsgBox "End Date must be entered"
            dlpTo(X).SetFocus
            Exit Function
        End If
    End If

    If Len(txtSeq(X%)) < 1 Then Exit For  ' missed one
    intRangesSet% = intRangesSet% + 1
Next X%
If intRangesSet% = 0 Then
    MsgBox "At least one Pay Period must be set"
    txtSeq(0).SetFocus
    Exit Function
End If


chkPayPeriod = True

Exit Function

chkPayPeriod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEntitle", "HRBENFT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Sub cmdCancel_Click()

fglbNew = False

Data1.Refresh

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Call Display_Value

vbxTrueGrid.SetFocus

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Sub cmdDelete_Click()
Dim SQLQ, Msg, a%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "The Pay Period?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Call getWSQLQ("C")

SQLQ = "DELETE FROM HR_PAYPERIOD WHERE " & fglbVSQLQ

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

Data1.Refresh

Call Display_Value

End Sub

Sub cmdModify_Click()
ODIV = clpDiv.Text
ODept = clpDept.Text
oOrg = clpCode(0).Text
oYear = txtYear.Text
OLoc = clpCode(4).Text
OSection = clpCode(3).Text
oPayP = clpPAYP.Text
Actn = "M"
End Sub

Sub cmdNew_Click()
Dim X

For X = 0 To txtSeq.count - 1
    txtSeq(X) = ""
    dlpFrom(X) = ""
    dlpTo(X) = ""
    chkUploaded(X) = False
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
clpPAYP.Text = ""
clpCode(3).Text = ""
clpCode(4).Text = ""

Actn = "A"
fglbNew = True

Call SET_UP_MODE

'clpDiv.SetFocus

End Sub

Sub cmdOK_Click()
Dim X%, Y%, xUnion, xPT, SQLQ, SQLQW
Dim xStr
Dim rsVE As New ADODB.Recordset
Dim rsVT As New ADODB.Recordset
Dim glbiOneWhere As Boolean
Dim strMsg As String
Dim Response%

If Not chkPayPeriod() Then Exit Sub

For X% = 0 To txtSeq.count - 1
    If Not IsNumeric(txtSeq(X)) Then Exit For
Next

'Ticket #29617 - Mississaugas of Scugog Island First Nation
'Trying to Close Pay Period from here?
If glbCompSerial = "S/N - 2485W" Then
    If Closing_PayPeriod Then
        Response% = MsgBox("Closing a Pay Period on this screen will not perform any entitlement updates." & Chr(10) & Chr(10) & "Are you sure you wish to proceed?", vbQuestion + vbYesNo, "Closing a Pay Period")
        If Response% = IDNO Then
            MsgBox "No changes have been saved.", vbInformation, "Save Aborted"
            Data1.Refresh
            Exit Sub
        End If
    End If
End If

'Ticket #23151 - Check to see if any of the Pay Period has changed. If changed then check if there is any Timesheet
'recorded for that Pay Period. If Timesheet found then don't allow user to save the changes. Pay Periods cannot be
'changed for which Timesheet records exists.
If Actn = "M" Then
    xPPNoChanged = ""
    
    'Do not allow to save changes as TS records were found for the Pay Period being changed.
    If Not AllowSaveChanges Then
        strMsg = "You cannot save the changes made to Pay Period #" & xPPNoChanged & ". There are Timesheet(s) records found for this Pay Period which may become unmanageable or will be detached from the Timesheet Module." & vbCrLf & vbCrLf
        strMsg = strMsg & "Please refrain from making changes to the existing/current Pay Periods."
        strMsg = strMsg & "You can only make changes to the Pay Periods which does not have any associated Timesheet records. But make sure you first stop access to Timesheet Module and info:HR before making these changes."
        MsgBox strMsg, vbExclamation, "Cannot save changes to these Pay Periods"
        
        Exit Sub
    End If
End If

If Actn = "M" Then
    Call getWSQLQ("O")
    SQLQ = "DELETE FROM HR_PAYPERIOD WHERE " & fglbVSQLQ
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    Call getWSQLQ("C")
    SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE " & fglbVSQLQ
    rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsVT.EOF Then
        MsgBox "You cannot add duplicate record"
        clpDiv.SetFocus
        Exit Sub
    End If
End If

gdbAdoIhr001.BeginTrans
SQLQ = "SELECT * FROM HR_PAYPERIOD"
rsVE.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
For X% = 0 To txtSeq.count - 1
    If Len(txtSeq(X%)) > 0 Then
        rsVE.AddNew
        rsVE("PP_YEAR") = txtYear.Text
        rsVE("PP_NBR") = txtSeq(X)
        rsVE("PP_ORG_TABL") = "EDOR"
        rsVE("PP_ORG") = clpCode(0).Text
        rsVE("PP_PAYP") = clpPAYP.Text
        rsVE("PP_DIV") = clpDiv.Text
        rsVE("PP_DEPTNO") = clpDept.Text
        rsVE("PP_SECTION") = clpCode(3).Text
        rsVE("PP_LOC") = clpCode(4).Text
        rsVE("PP_START") = dlpFrom(X%)
        rsVE("PP_END") = dlpTo(X%)
        rsVE("PP_UPLOADED") = chkUploaded(X%) <> 0
        rsVE.Update
                
        'Leeds and Grenville - Ticket #19441 - OT/CT Weekly Adjustment for the closed Pay Period
        If glbCompSerial = "S/N - 2233W" And rsVE("PP_UPLOADED") Then
            Call LeedsGrenville_OTAdjustments(rsVE("PP_START"), rsVE("PP_END"))
        End If
    End If
Next
rsVE.Close
gdbAdoIhr001.CommitTrans

'If Not glbSQL and not glboracle Then Call Pause(0.5)
Data1.Refresh

fglbNew = False

Call Display_Value

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Pay Period Master Report"
Call setRptLabel(Me, 0) '1)
Me.vbxCrystal.Connect = RptODBC_SQL
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgpayperiod.rpt"

SQLQ = "(1=1) "
If Len(txtYear.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_YEAR) = " & txtYear.Text
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_DEPTNO} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_ORG} = '" & clpCode(0).Text & "'"
If Len(clpPAYP.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_PAYP} = '" & clpPAYP.Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_LOC} = '" & clpCode(4).Text & "'"

Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Pay Period Master Report"
Call setRptLabel(Me, 0) '1)

Me.vbxCrystal.Connect = RptODBC_SQL

Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgpayperiod.rpt"

SQLQ = "(1=1) "
If Len(txtYear.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_YEAR} = " & txtYear.Text
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_DIV} = '" & clpDiv.Text & "'"
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_DEPTNO} = '" & clpDept.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_ORG} = '" & clpCode(0).Text & "'"
If Len(clpPAYP.Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_PAYP} = '" & clpPAYP.Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_SECTION} = '" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND {HR_PAYPERIOD.PP_LOC} = '" & clpCode(4).Text & "'"

Me.vbxCrystal.SelectionFormula = SQLQ
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Private Sub cmdPrintAll_Click()
Dim RHeading As String, xReport, X%
Dim SQLQ
Dim dtYYY%, dtMM%, dtDD%
cmdPrintAll.Enabled = False

Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Pay Period Master Report"
Call setRptLabel(Me, 0) '1)
Me.vbxCrystal.Connect = RptODBC_SQL

Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgPayperiod.rpt"
Me.vbxCrystal.Action = 1

cmdPrintAll.Enabled = True
End Sub

Private Sub comNbrOfPayPeriod_Click()
Call AutoSetPayPeriod
End Sub

Private Sub dlpFrom_GotFocus(Index As Integer)
If Index = 0 Then
    oldStart = dlpFrom(0)
End If
End Sub

Private Sub dlpFrom_LostFocus(Index As Integer)
If Index = 0 Then Call AutoSetPayPeriod
End Sub

Private Sub AutoSetPayPeriod()
Dim X

If Not fglbNew Then Exit Sub
If Not IsDate(dlpFrom(0)) Then Exit Sub
If Val(comNbrOfPayPeriod.Text) = 26 Or Val(comNbrOfPayPeriod.Text) = 27 Then
    For X = 0 To txtSeq.count - 1
        If X <> 0 Then
            txtSeq(X) = ""
            dlpFrom(X) = ""
            dlpTo(X) = ""
            chkUploaded(X) = False
        End If
        If X < Val(comNbrOfPayPeriod.Text) Then
            txtSeq(X) = X + 1
            dlpFrom(X) = DateAdd("d", 14 * X, dlpFrom(0))
            dlpTo(X) = DateAdd("d", 14 * X + 13, dlpFrom(0))
        End If
    Next
End If
If Val(comNbrOfPayPeriod.Text) = 52 Or Val(comNbrOfPayPeriod.Text) = 53 Then
    For X = 0 To txtSeq.count - 1
        If X <> 0 Then
            txtSeq(X) = ""
            dlpFrom(X) = ""
            dlpTo(X) = ""
            chkUploaded(X) = False
        End If
        If X < Val(comNbrOfPayPeriod.Text) Then
            txtSeq(X) = X + 1
            dlpFrom(X) = DateAdd("d", 7 * X, dlpFrom(0))
            dlpTo(X) = DateAdd("d", 7 * X + 6, dlpFrom(0))
        End If
    Next
End If

If Val(comNbrOfPayPeriod.Text) = 12 Then
    For X = 0 To txtSeq.count - 1
        If X <> 0 Then
            txtSeq(X) = ""
            dlpFrom(X) = ""
            dlpTo(X) = ""
            chkUploaded(X) = False
        End If
        If X < Val(comNbrOfPayPeriod.Text) Then
            txtSeq(X) = X + 1
            dlpFrom(X) = DateAdd("m", X, dlpFrom(0))
            dlpTo(X) = DateAdd("d", -1, DateAdd("m", X + 1, dlpFrom(0)))
        End If
    Next
End If
If Val(comNbrOfPayPeriod.Text) = 24 Then
    For X = 0 To txtSeq.count - 1
        If X <> 0 Then
            txtSeq(X) = ""
            dlpFrom(X) = ""
            dlpTo(X) = ""
            chkUploaded(X) = False
        End If
        If X < Val(comNbrOfPayPeriod.Text) Then
            txtSeq(X) = X + 1
            If Int(X / 2) = X / 2 Then
                dlpFrom(X) = DateAdd("m", Int(X / 2), dlpFrom(0))
                dlpTo(X) = DateAdd("d", 14, dlpFrom(X))
            Else
                dlpFrom(X) = DateAdd("d", 1, dlpTo(X - 1))
                dlpTo(X) = DateAdd("d", -1, DateAdd("m", Int(X / 2) + 1, dlpFrom(0)))
            End If
        End If
    Next
End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Call INI_Controls(Me)
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMPAYPERIODMASTER"

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ

FlagRefresh = False

comNbrOfPayPeriod.AddItem "12"
comNbrOfPayPeriod.AddItem "24"
comNbrOfPayPeriod.AddItem "26"
comNbrOfPayPeriod.AddItem "27"
comNbrOfPayPeriod.AddItem "52"
comNbrOfPayPeriod.AddItem "53"


Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT DISTINCT PP_YEAR,PP_DIV,PP_DEPTNO,PP_ORG,PP_LOC,PP_SECTION,PP_PAYP,PP_YEAR AS PPYEAR FROM HR_PAYPERIOD "
If glbDIVCount = 1 And glbLinamar Then
    SQLQ = SQLQ & " WHERE PP_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
End If

'Ticket #28359
SQLQ = SQLQ & " ORDER BY PP_YEAR DESC"

Data1.RecordSource = SQLQ
Data1.Refresh

Screen.MousePointer = HOURGLASS
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(1).Caption = lStr(vbxTrueGrid.Columns(1).Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)

Call setRptCaption(Me)
Call setCaption(lblTitle(12))

'Ticket #29861 Franks 03/28/2017 - WFC setup based on Pay Period "SM" or "W", not based on Plant, so remove the following logic
''If glbWFC Then
''    lblSection.FontBold = True
''End If

Screen.MousePointer = DEFAULT

Call INI_Controls(Me)

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
'Dim Keepfocus As Boolean
'If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
'Keepfocus = Not isUpdated(Me)
'Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Resize()
If Me.Height >= 3750 + VacFram.Height + panControls.Height + 230 Then
    scrControl.Value = 0
    VacFram.Top = 3960
    scrControl.Visible = False
    Exit Sub
End If
scrControl.Visible = True
scrControl.Max = VacFram.Height + panControls.Height + 3750 + 550 - Me.Height '250 - Me.Height
scrControl.Left = Me.Width - scrControl.Width - 260 '120
If Me.Height - scrControl.Top - panControls.Height - 300 > 0 Then
    scrControl.Height = Me.Height - scrControl.Top - panControls.Height - 300
Else
    scrControl.Height = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."

Set frmUEntitle = Nothing  'carmen apr 2000
End Sub

Private Sub scrControl_Change()
VacFram.Top = 3960 - scrControl.Value
End Sub

Sub ST_UPD_MODE(TF As Boolean)
Dim X, FT
FT = Not TF
For X = 0 To txtSeq.count - 1
    txtSeq(X).Enabled = TF
    dlpFrom(X).Enabled = TF
    dlpTo(X).Enabled = TF
    chkUploaded(X).Enabled = TF
Next

clpDiv.Enabled = TF
clpDept.Enabled = TF
clpCode(0).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
lblNbrOfPayPeriod.Visible = fglbNew
comNbrOfPayPeriod.Visible = fglbNew

End Sub

Sub Display_Value()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim X
For X = 0 To txtSeq.count - 1
    txtSeq(X) = ""
    dlpFrom(X) = ""
    dlpTo(X) = ""
    chkUploaded(X) = False
Next
clpDiv.Text = ""
clpDept.Text = ""
clpCode(0).Text = ""
clpPAYP.Text = ""
clpCode(3).Text = ""
clpCode(4).Text = ""
txtYear = ""
If Not Data1.Recordset.EOF Then
    SQLQ = "SELECT * FROM HR_PAYPERIOD "
    If IsNull(Data1.Recordset("PP_YEAR")) Then
        SQLQ = SQLQ & " WHERE PP_YEAR IS NULL"
    Else
        SQLQ = SQLQ & " WHERE PP_YEAR = " & Data1.Recordset("PP_YEAR")
    End If
    If IsNull(Data1.Recordset("PP_DIV")) Then
        SQLQ = SQLQ & " AND PP_DIV IS NULL"
    Else
        SQLQ = SQLQ & " AND PP_DIV = '" & Data1.Recordset("PP_DIV") & "'"
    End If
    If IsNull(Data1.Recordset("PP_DEPTNO")) Then
        SQLQ = SQLQ & " AND PP_DEPTNO IS NULL"
    Else
        SQLQ = SQLQ & " AND PP_DEPTNO = '" & Data1.Recordset("PP_DEPTNO") & "'"
    End If
    If IsNull(Data1.Recordset("PP_ORG")) Then
        SQLQ = SQLQ & " AND PP_ORG IS NULL"
    Else
        SQLQ = SQLQ & " AND PP_ORG = '" & Data1.Recordset("PP_ORG") & "'"
    End If
    If IsNull(Data1.Recordset("PP_LOC")) Then
        SQLQ = SQLQ & " AND PP_LOC IS NULL"
    Else
        SQLQ = SQLQ & " AND PP_LOC = '" & Data1.Recordset("PP_LOC") & "'"
    End If
    If IsNull(Data1.Recordset("PP_SECTION")) Then
        SQLQ = SQLQ & " AND PP_SECTION IS NULL"
    Else
        SQLQ = SQLQ & " AND PP_SECTION = '" & Data1.Recordset("PP_SECTION") & "'"
    End If
    If IsNull(Data1.Recordset("PP_PAYP")) Then
        SQLQ = SQLQ & " AND PP_PAYP IS NULL"
    Else
        SQLQ = SQLQ & " AND PP_PAYP = '" & Data1.Recordset("PP_PAYP") & "'"
    End If
    
    SQLQ = SQLQ & " Order By PP_YEAR,PP_DIV,PP_DEPTNO,PP_ORG, PP_LOC,PP_SECTION,PP_NBR "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not IsNull(Data1.Recordset("PP_YEAR")) Then txtYear.Text = Data1.Recordset("PP_YEAR")
    If Not IsNull(Data1.Recordset("PP_DIV")) Then clpDiv.Text = Data1.Recordset("PP_DIV")
    If Not IsNull(Data1.Recordset("PP_DEPTNO")) Then clpDept.Text = Data1.Recordset("PP_DEPTNO")
    If Not IsNull(Data1.Recordset("PP_ORG")) Then clpCode(0).Text = Data1.Recordset("PP_ORG")
    If Not IsNull(Data1.Recordset("PP_LOC")) Then clpCode(4).Text = Data1.Recordset("PP_LOC")
    If Not IsNull(Data1.Recordset("PP_SECTION")) Then clpCode(3).Text = Data1.Recordset("PP_SECTION")
    If Not IsNull(Data1.Recordset("PP_PAYP")) Then clpPAYP.Text = Data1.Recordset("PP_PAYP")
    
    Do While Not rsVE.EOF
        xOrder = rsVE("PP_NBR")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 53) Then
            If Not IsNull(rsVE("PP_NBR")) Then txtSeq(nOrder) = rsVE("PP_NBR")
            If Not IsNull(rsVE("PP_START")) Then dlpFrom(nOrder) = rsVE("PP_START")
            If Not IsNull(rsVE("PP_END")) Then dlpTo(nOrder) = rsVE("PP_END")
            If rsVE("PP_UPLOADED") <> 0 Then
                chkUploaded(nOrder) = 1
            Else
                chkUploaded(nOrder) = 0
            End If
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close
End If

Call SET_UP_MODE

Call cmdModify_Click

End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
          
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT DISTINCT PP_YEAR,PP_DIV,PP_DEPTNO,PP_ORG,PP_LOC,PP_SECTION,PP_PAYP,PP_YEAR AS PPYEAR FROM HR_PAYPERIOD "
    If glbDIVCount = 1 And glbLinamar Then
        SQLQ = SQLQ & " WHERE PP_DIV IN (select DIV from HR_DIVISION WHERE " & glbSeleDiv & ")"
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub

Private Sub getWSQLQ(xType)
Dim xDiv, xDept, xORG, xPAYP
Dim xLoc, xSection, xYear

fglbESQLQ = glbSeleDeptUn

If Len(txtYear.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_YEAR = '" & txtYear.Text & "' "
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & clpDept.Text & "' "
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & clpCode(0).Text & "' "
If Len(clpPAYP.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PAYP = '" & clpPAYP.Text & "' "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & clpCode(3).Text & "' "
If Len(clpCode(4).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & clpCode(4).Text & "' "
If xType = "" Then Exit Sub

If xType = "O" Then
    xYear = oYear
    xDiv = ODIV
    xDept = ODept
    xORG = oOrg
    xPAYP = oPayP
    xLoc = OLoc
    xSection = OSection
Else
    xYear = txtYear.Text
    xDiv = clpDiv.Text
    xDept = clpDept.Text
    xORG = clpCode(0).Text
    xPAYP = clpPAYP.Text
    xLoc = clpCode(4).Text
    xSection = clpCode(3).Text
End If
If Len(xYear) = 0 Then
    fglbVSQLQ = " (PP_YEAR IS NULL OR PP_YEAR =0)"
Else
    fglbVSQLQ = " PP_YEAR= " & xYear
End If
If Len(xDiv) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PP_DIV IS NULL OR PP_DIV='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND PP_DIV = '" & xDiv & "'"
End If
If Len(xDept) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PP_DEPTNO IS NULL OR PP_DEPTNO='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND PP_DEPTNO = '" & xDept & "'"
End If
If Len(xORG) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PP_ORG IS NULL OR PP_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PP_ORG = '" & xORG & "'"
End If
If Len(xPAYP) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PP_PAYP IS NULL OR PP_PAYP='')"
Else
    fglbVSQLQ = fglbVSQLQ & " AND PP_PAYP = '" & xPAYP & "'"
End If

If Len(xLoc) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PP_LOC IS NULL OR PP_LOC='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PP_LOC = '" & xLoc & "'"
End If
If Len(xSection) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (PP_SECTION IS NULL OR PP_SECTION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND PP_SECTION = '" & xSection & "'"
End If

End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
    cmdPrintAll.Enabled = False
ElseIf Me.Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
    cmdPrintAll.Enabled = True
Else
    UpdateState = OPENING
    TF = True
    cmdPrintAll.Enabled = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

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
UpdateRight = gSec_Upd_PayPeriod_Master
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

'Ticket #23275 - OT Logic Revised.
'New Routine containing the Revised logic. Not overwriting this to be on the safe side
'Private Sub LeedsGrenville_OTAdjustments(xPayStartDate, xPayEndDate)
'
''Forfeit OT Hours
''Create CTEX - Overtime Expired code if not already existing
''1. Retrieve OT records from HR_ATTENDANCE where AD_BANKHRS_EXP < Pay Period End Date
''2. For each record retrieved,
''   - remove AD_BANKHRS_EXP value, Frozen (AD_INDICATOR (Incentive)), LDate(NOW), LUser, LTime,
''   - Create a CTEX record for the same hours for the Pay Period End Date, Frozen(AD_INDICATOR (Incentive)).
'
''OT Hours Adjustments
''Create CTRV - Reversed for weekly adjustment code if not already existing
''Create CTPD - Overtime Paid code if not already existing
''1. Retrieve OT Records from HR_ATTENDANCE where AD_BANKHRS_EXP is not blank or null, and
''   Not Frozen - for the Pay Period just marked as closed.
''2. Calculate the Total Hours Worked for the week, i.e. OT hours retrieved + Hours/Week (35)
''3. Create Reversal Entry (CTRV) for total OT retrieved for the week as of Pay End Date
''4. For each record retrieved, make the following adjustment:
''   - >35 < 40  (4hrs - OT * 1)
''   - >=40 < 44 (5hrs - OT15 * 1.5)
''   - >=44      (CTPD * 1.5)
''  If Total Hours Worked - 35 >= 4 then OT Hours * 1
''       - Create an OT Record for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
''  If Total Hours Worked - 35 > 4 and <= 8 then 4 OT Hours * 1, (OT Hours - 4) * 1.5
''       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
''       - Create an OT15 Record for (OT Hours - 4) * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
''  If Total Hours Worked - 35 > 8 then 4 OT Hours * 1, 4 OT Hours * 1.5, (OT Hours - 8) * 1.5
''       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
''       - Create an OT15 Record for 4 hrs * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
''       - Create an OTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen, Null Expiry Date, Consumed Fully
''           - this is an Overtime Paid adjustment entry for which an automatic CTPD will be done.
''       - Create an CTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen
''           - for OTPD - Overtime Paid Out. This will balance out the OT Outstanding otherwise there will
''             always be an overused CT balance.
''  Update the original OT record retrieved
''       - Remove the Expiry Date
''       - Freeze it
'
''Adjusted OT record Consumed by CT
''   - Retrieve CT records which are not frozen and not fully adjusted.
''   - Retrieve OT records which are Frozen, Not fully consumed and with Expiry Date
''       - Update AD_CONSUMED with CT hours consumed from this OT record
''           - If fully consumed, remove Expiry Date and Add Comments as used OT
''       - Move through each OT record until CT is fully adjusted
''           - Update AD_Consumed with hours adjusted
'
''Ticket #21411 - OTRE - Overtime Reinstated. These are expired hours that the user will be re-entering
''back so the employees can take it. They don't want employees to loose out hours because they are still
''trying to understand the logic of OT expired and also they are not really aware of how many hours they
''have.
'
'    Dim SQLQ As String
'    Dim rsAttend As New ADODB.Recordset
'    Dim rsAddAttend As New ADODB.Recordset
'    Dim rsConsAttend As New ADODB.Recordset
'    Dim OTHours As Double
'    Dim HrsWorked As Double
'    Dim StdHrsWeek As Double
'    Dim OT15Hrs As Double
'    Dim CTPDHrs As Double
'    Dim AdjHrs As Double
'    Dim OTBal As Double
'    Dim CTBal As Double
'    Dim xExpiryDate
'
'
'    OTHours = 0
'    AdjHrs = 0
'    HrsWorked = 0
'    StdHrsWeek = 35
'    OT15Hrs = 40
'    CTPDHrs = 44
'
'    'Forfeit Hours Begin --------------------------------------------------------------------------------
'    'Create CTEX - Overtime Expired code if not already existing
'    Call CreateTableMasterCode("ADRE", "CTEX", "Overtime Expired")
'
'    'Retrieve OT Attendance records which have expired and Frozen
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'    SQLQ = SQLQ & " AND AD_BANKHRS_EXP <" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_REASON LIKE ('OT%')"
'    SQLQ = SQLQ & " AND AD_INDICATOR = 1"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_BANKHRS_EXP, AD_DOA"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    Do While Not rsAttend.EOF
'        'Remove the expiry date from these records
'        xExpiryDate = rsAttend("AD_BANKHRS_EXP")
'        rsAttend("AD_BANKHRS_EXP") = Null
'        'rsAttend("AD_INDICATOR") = 1  'Freezing record
'        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & Round((rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))), 2) & " Overtime Hours Expired and " & IIf(IsNull(rsAttend("AD_CONSUMED")), 0, Round(rsAttend("AD_CONSUMED"), 2)) & " Overtime Hours Consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_LDATE") = Date
'        rsAttend("AD_LTIME") = Time$
'        rsAttend("AD_LUSER") = glbUserID
'        rsAttend.Update
'
'        'Create a CTEX record for the same hours
'        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
'        rsAddAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        rsAddAttend.AddNew
'        rsAddAttend("AD_COMPNO") = "001"
'        rsAddAttend("AD_EMPNBR") = rsAttend("AD_EMPNBR")
'        rsAddAttend("AD_SALARY") = rsAttend("AD_SALARY")
'        rsAddAttend("AD_SALCD") = rsAttend("AD_SALCD")
'        rsAddAttend("AD_JOB") = rsAttend("AD_JOB")
'        rsAddAttend("AD_DHRS") = rsAttend("AD_DHRS")
'        rsAddAttend("AD_WHRS") = rsAttend("AD_WHRS")
'        rsAddAttend("AD_SUPER") = rsAttend("AD_SUPER")
'        rsAddAttend("AD_PAYROLL_ID") = rsAttend("AD_PAYROLL_ID")
'        rsAddAttend("AD_SHIFT") = rsAttend("AD_SHIFT")
'        rsAddAttend("AD_GLNO") = rsAttend("AD_GLNO")
'        rsAddAttend("AD_ORG") = rsAttend("AD_ORG")
'        rsAddAttend("AD_DOA") = xPayEndDate
'        rsAddAttend("AD_REASON") = "CTEX"
'        rsAddAttend("AD_HRS") = rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))
'        rsAddAttend("AD_CONSUMED") = rsAddAttend("AD_HRS")
'        rsAddAttend("AD_COMM") = "Weekly Forfeited Hours for the Attendance record '" & rsAttend("AD_REASON") & "' dated '" & Format(rsAttend("AD_DOA"), "Short Date") & "' with Expiry Date of '" & Format(xExpiryDate, "Short Date") & "." & vbCrLf & Round((rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))), 2) & " Overtime Hour(s) Expired and " & IIf(IsNull(rsAttend("AD_CONSUMED")), 0, Round(rsAttend("AD_CONSUMED"), 2)) & " Overtime Hour(s) Consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAddAttend("AD_INDICATOR") = 1  'Freezing record
'        rsAddAttend("AD_LUSER") = glbUserID
'        rsAddAttend("AD_LDATE") = Date
'        rsAddAttend("AD_LTIME") = Time$
'        rsAddAttend("AD_SOURCE") = "IHRFOR"
'        rsAddAttend.Update
'        rsAddAttend.Close
'        Set rsAddAttend = Nothing
'
'        rsAttend.MoveNext
'    Loop
'    rsAttend.Close
'    Set rsAttend = Nothing
'    'Forfeit Hours End ----------------------------------------------------------------------------------
'
'    'OT Hours Adjustments Begin -------------------------------------------------------------------------
'    'Create CTRV - OT Reversed code, CTPD - Overtime Paid code, if not already existing
'    Call CreateTableMasterCode("ADRE", "CTRV", "Overtime Reversed")
'    Call CreateTableMasterCode("ADRE", "OTPD", "Overtime Earned for Pay Out")
'    Call CreateTableMasterCode("ADRE", "CTPD", "Overtime Paid Out")
'    Call CreateTableMasterCode("ADRE", "OT15", "Overtime at Time and One-Half")
'
'    'Ticket #21197 Begin ------------------------------------------------------------------------
'    'Adjust/Consume the OTs and CTs of the Pay Week before Reversing the original entry
'    'Retrieve CT Attendance records for the Pay Period just closed where Not Frozen
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE "
'    SQLQ = SQLQ & " AD_DOA >=" & Date_SQL(xPayStartDate)
'    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
'    SQLQ = SQLQ & " AND AD_REASON = 'CT'"
'    SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR ASC, AD_DOA ASC"
'    rsConsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'    Do While Not rsConsAttend.EOF
'
'        'Retrieve OT Attendance for the Pay Period not Consumed fully, not expired
'        'and not Frozen
'        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'        SQLQ = SQLQ & " AND AD_EMPNBR = " & rsConsAttend("AD_EMPNBR")
'        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
'        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'        SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
'        SQLQ = SQLQ & " AND AD_REASON LIKE 'OT%'"
'        SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
'        SQLQ = SQLQ & " ORDER BY AD_DOA ASC"
'        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'        Do While Not rsAttend.EOF
'            'If fully consumed - Remove Expiry Date, add Comments
'
'            'Check if there is enough hours OT (Ad_hrs  - Consumed) left to be
'            'consumed by CT (Ad_Hrs - Consumed) left
'            OTBal = 0
'            CTBal = 0
'            OTBal = (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")))
'            CTBal = rsConsAttend("AD_HRS") - IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED"))
'            If OTBal >= CTBal Then
'                'Enough OT Hours left to consume CT Bal
'                'Update OT Consumed with all of CT Bal - some OT will still be left
'                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + CTBal
'
'                'Update CT Consumed with CT Bal - nothing will be left
'                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + CTBal
'
'                'Update other fields of both these record sets
'                'OT recordset
'                If OTBal = CTBal Then
'                    'Since all OT Hours has been used up, remove expiry date and update Comments as consumed
'                    rsAttend("AD_BANKHRS_EXP") = Null
'                    rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All of original OT hours consumed within the week of " & Format(xPayStartDate, "mmm dd, yyyy") & " - " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'                    rsAttend("AD_INDICATOR") = 1  'Freezing record
'                End If
'                rsAttend("AD_LDATE") = Date
'                rsAttend("AD_LTIME") = Time$
'                rsAttend("AD_LUSER") = glbUserID
'                rsAttend.Update
'
'                'CT recordset
'                rsConsAttend("AD_INDICATOR") = 1  'Freezing record
'                rsConsAttend("AD_LDATE") = Date
'                rsConsAttend("AD_LTIME") = Time$
'                rsConsAttend("AD_LUSER") = glbUserID
'                rsConsAttend.Update
'
'                'All CT hours have balanced up with OT - Move to next CT record
'                Exit Do
'            Else
'                'OT Hours left to consume is less that CT Bal to consume
'                'Update CT Consume with OT hours left to be consumed (OTBal)
'                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + OTBal
'
'                'Update OT Consume with OT hours left to be consumed (OTBal) = all OT hours used up
'                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + OTBal
'
'                'Update OT recordset
'                'Since all OT hours has been used up, remove expiry date, update Comments as consumed
'                rsAttend("AD_BANKHRS_EXP") = Null
'                rsAttend("AD_INDICATOR") = 1  'Freezing record
'                rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All of original OT hours consumed within the week of " & Format(xPayStartDate, "mmm dd, yyyy") & " - " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'                rsAttend("AD_LDATE") = Date
'                rsAttend("AD_LTIME") = Time$
'                rsAttend("AD_LUSER") = glbUserID
'                rsAttend.Update
'
'                'Update CT recordset
'                'rsConsAttend("AD_INDICATOR") = 1  'Freezing record
'                rsConsAttend("AD_LDATE") = Date
'                rsConsAttend("AD_LTIME") = Time$
'                rsConsAttend("AD_LUSER") = glbUserID
'                rsConsAttend.Update
'
'            End If
'
'            rsAttend.MoveNext
'        Loop
'        rsAttend.Close
'        Set rsAttend = Nothing
'
'        rsConsAttend.MoveNext
'    Loop
'    rsConsAttend.Close
'    Set rsConsAttend = Nothing
'    'Ticket #21197 End --------------------------------------------------------------------------
'
'
'    'OT Reversal Entry
'    'Retrieve OT Attendance records for the Pay Period just closed
'    'where AD_BANKHRS_EXP is not blank or null, and Not Frozen
'    SQLQ = "SELECT SUM(CASE WHEN AD_HRS IS NULL THEN 0 ELSE AD_HRS END) - SUM(CASE WHEN AD_CONSUMED IS NULL THEN 0 ELSE AD_CONSUMED END) AS TOT_OT, AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
'    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
'    SQLQ = SQLQ & " AND AD_REASON = 'OT'"
'    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'    Do While Not rsAttend.EOF
'        'Calculate the Total Hours Worked for the week, i.e. OT hours retrieved + Hours/Week (35)
'        OTHours = IIf(IsNull(rsAttend("TOT_OT")), 0, rsAttend("TOT_OT"))
'        HrsWorked = StdHrsWeek + OTHours
'
'        'Create CTRV - OT Reversal Entry
'        Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "CTRV", OTHours, "", "OT Reversal entry for the adjustment.", OTHours)
'
'
'        'Create the OT Adjustment entries
'        '   - >35 < 40  (4hrs - OT * 1)
'        '   - >=40 < 44 (5hrs - OT15 * 1.5)
'        '   - >=44      (CTPD * 1.5)
'        '  If Total Hours Worked - 35 >= 4 then OT Hours * 1
'        '       - Create an OT Record for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'        '  If Total Hours Worked - 35 > 4 and <= 8 then 4 OT Hours * 1, (OT Hours - 4) * 1.5
'        '       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'        '       - Create an OT15 Record for (OT Hours - 4) * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'        '  If Total Hours Worked - 35 > 8 then 4 OT Hours * 1, 4 OT Hours * 1.5, (OT Hours - 8) * 1.5
'        '       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'        '       - Create an OT15 Record for 4 hrs * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'        '       - Create an CTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen
'
'        If (HrsWorked - StdHrsWeek) <= 4 Then
'        'If (HrsWorked > StdHrsWeek) And (HrsWorked < OT15Hrs) Then     '> 35 < 40
'            'Ajusted OT Hours
'            AdjHrs = HrsWorked - StdHrsWeek
'
'            'Create an OT Record for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            If AdjHrs > 0 Then
'                Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'            End If
'        End If
'
'        If (HrsWorked - StdHrsWeek) > 4 And (HrsWorked - StdHrsWeek) <= 8 Then
'        'If (HrsWorked >= OT15Hrs) And (HrsWorked < CTPDHrs) Then    '>=40 < 44
'            'Ajusted OT Hours for OT
'            AdjHrs = 4
'
'            'Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'
'            'Ajusted OT Hours for OT15
'            AdjHrs = (HrsWorked - (OT15Hrs - 1)) * 1.5
'
'            'Create an OT15 Record for (OT Hours - 4) * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT15", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'        End If
'
'        If (HrsWorked - StdHrsWeek) > 8 Then
'        'If (HrsWorked >= CTPDHrs) Then    '>=44
'            'Ajusted OT Hours for OT
'            AdjHrs = 4
'
'            'Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'
'
'            'Ajusted OT Hours for OT15
'            AdjHrs = 4 * 1.5
'
'            'Create an OT15 Record for 4 hrs * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT15", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'
'
'            'Ajusted OT Hours for CTPD
'            AdjHrs = (HrsWorked - (CTPDHrs - 1)) * 1.5
'
'            'Create an OTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen, Null Expiry Date, Fully Consumed
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OTPD", AdjHrs, "", "Overtime Earned for Pay Out Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."), AdjHrs)
'
'            'Create an CTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "CTPD", AdjHrs, "", "Paid Out Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."), AdjHrs)
'        End If
'
'        rsAttend.MoveNext
'    Loop
'    rsAttend.Close
'    Set rsAttend = Nothing
'
'    'Freeze the original OT entries for the Pay Period just closed
'    'Retrieve OT Attendance records for the Pay Period just closed
'    'where AD_BANKHRS_EXP is not blank or null, and Not Frozen
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
'    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
'    SQLQ = SQLQ & " AND AD_REASON = 'OT'"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    Do While Not rsAttend.EOF
'        'Remove the expiry date from these records and freeze it
'        rsAttend("AD_BANKHRS_EXP") = Null
'        rsAttend("AD_INDICATOR") = 1  'Freezing record
'        rsAttend("AD_CONSUMED") = rsAttend("AD_HRS")
'        'rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "Original OT entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))) & " hour(s) from the original " & rsAttend("AD_HRS") & " OT hours entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_LDATE") = Date
'        rsAttend("AD_LTIME") = Time$
'        rsAttend("AD_LUSER") = glbUserID
'        rsAttend.Update
'
'        rsAttend.MoveNext
'    Loop
'    rsAttend.Close
'    Set rsAttend = Nothing
'
'    'Ticket #21411 - Since they are giving back the expired hours - need to freeze it so it can be consumed.
'    'Freeze the OTRE entries for the Pay Period just closed
'    'Retrieve OTRE Attendance records for the Pay Period just closed
'    'where AD_BANKHRS_EXP is not blank or null, and Not Frozen
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
'    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
'    SQLQ = SQLQ & " AND AD_REASON = 'OTRE'"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    Do While Not rsAttend.EOF
'        'Remove the expiry date from these records and freeze it
'        'rsAttend("AD_BANKHRS_EXP") = Null
'        rsAttend("AD_INDICATOR") = 1  'Freezing record
'        'rsAttend("AD_CONSUMED") = rsAttend("AD_HRS")
'        'rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "Original OT entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & rsAttend("AD_HRS") & " OTRE hours entry frozen as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_LDATE") = Date
'        rsAttend("AD_LTIME") = Time$
'        rsAttend("AD_LUSER") = glbUserID
'        rsAttend.Update
'
'        rsAttend.MoveNext
'    Loop
'    rsAttend.Close
'    Set rsAttend = Nothing
'
'    'Freeze the CT entries for the Pay Period just closed
'    'Retrieve CT Attendance records for the Pay Period just closed
'    'where Not Frozen
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE "
'    SQLQ = SQLQ & " AD_DOA >=" & Date_SQL(xPayStartDate)
'    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
'    SQLQ = SQLQ & " AND AD_REASON = 'CT'"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    Do While Not rsAttend.EOF
'        'Freeze it
'        'rsAttend("AD_BANKHRS_EXP") = Null
'        rsAttend("AD_INDICATOR") = 1  'Freezing record
'        'rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "Original CT entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))) & " hour(s) from the original " & rsAttend("AD_HRS") & " CT hours entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_LDATE") = Date
'        rsAttend("AD_LTIME") = Time$
'        rsAttend("AD_LUSER") = glbUserID
'        rsAttend.Update
'
'        rsAttend.MoveNext
'    Loop
'    rsAttend.Close
'    Set rsAttend = Nothing
'
'    'OT Hours Adjustments End ---------------------------------------------------------------------------
'
'    'Consumed Hours (Adjust CT against OTs as Consumed Hours) Begin --------------------------------------
'    'Retrieve CT Attendance records, not fully adjusted
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE "
'    'SQLQ = SQLQ & " AD_DOA >=" & Date_SQL(xPayStartDate)
'    SQLQ = SQLQ & " AD_DOA <=" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_INDICATOR = 1"
'    SQLQ = SQLQ & " AND AD_REASON = 'CT'"
'    SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR ASC, AD_DOA ASC"
'    rsConsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'    Do While Not rsConsAttend.EOF
'
'        'Retrieve OT Attendance not Consumed fully, not expired and Frozen
'        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'        SQLQ = SQLQ & " AND AD_EMPNBR = " & rsConsAttend("AD_EMPNBR")
'        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
'        SQLQ = SQLQ & " AND AD_INDICATOR = 1"
'        SQLQ = SQLQ & " AND AD_REASON LIKE 'OT%'"
'        SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
'        SQLQ = SQLQ & " ORDER BY AD_DOA ASC"
'        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
'        Do While Not rsAttend.EOF
'            'If fully consumed - Remove Expiry Date, add Comments
'
'            'Check if there is enough hours OT (Ad_hrs  - Consumed) left to be consumed by CT (Ad_Hrs - Consumed) left
'            OTBal = 0
'            CTBal = 0
'            OTBal = (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")))
'            CTBal = rsConsAttend("AD_HRS") - IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED"))
'            If OTBal >= CTBal Then
'                'Enough OT Hours left to consume CT Bal
'                'Update OT Consumed with all of CT Bal - some OT will still be left
'                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + CTBal
'
'                'Update CT Consumed with CT Bal - nothing will be left
'                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + CTBal
'
'                'Update other fields of both these record sets
'                'OT recordset
'                If OTBal = CTBal Then
'                    'Since all OT Hours has been used up, remove expiry date and update Comments as consumed
'                    rsAttend("AD_BANKHRS_EXP") = Null
'                    rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All OT hours consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'                End If
'                rsAttend("AD_LDATE") = Date
'                rsAttend("AD_LTIME") = Time$
'                rsAttend("AD_LUSER") = glbUserID
'                rsAttend.Update
'
'                'CT recordset
'                rsConsAttend("AD_INDICATOR") = 1  'Freezing record
'                rsConsAttend("AD_LDATE") = Date
'                rsConsAttend("AD_LTIME") = Time$
'                rsConsAttend("AD_LUSER") = glbUserID
'                rsConsAttend.Update
'
'                'All CT hours have balanced up with OT - Move to next CT record
'                Exit Do
'            Else
'                'OT Hours left to consume is less that CT Bal to consume
'                'Update CT Consume with OT hours left to be consumed (OTBal)
'                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + OTBal
'
'                'Update OT Consume with OT hours left to be consumed (OTBal) = all OT hours used up
'                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + OTBal
'
'                'Update OT recordset
'                'Since all OT hours has been used up, remove expiry date, update Comments as consumed
'                rsAttend("AD_BANKHRS_EXP") = Null
'                rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All OT hours consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'                rsAttend("AD_LDATE") = Date
'                rsAttend("AD_LTIME") = Time$
'                rsAttend("AD_LUSER") = glbUserID
'                rsAttend.Update
'
'                'Update CT recordset
'                rsConsAttend("AD_INDICATOR") = 1  'Freezing record
'                rsConsAttend("AD_LDATE") = Date
'                rsConsAttend("AD_LTIME") = Time$
'                rsConsAttend("AD_LUSER") = glbUserID
'                rsConsAttend.Update
'
'            End If
'
'            rsAttend.MoveNext
'        Loop
'        rsAttend.Close
'        Set rsAttend = Nothing
'
'        rsConsAttend.MoveNext
'    Loop
'    rsConsAttend.Close
'    Set rsConsAttend = Nothing
'    'Consumed Hours (Adjust CT against OTs as Consumed Hours) End ----------------------------------------
'
'
'End Sub

Private Sub LeedsGrenville_OTAdjustments(xPayStartDate, xPayEndDate)

'Ticket #23275 - OT Logic Revised.
'New Logic:
'>35 < 37.5 (OT * 1)
'>=37.5 (OT15 * 1.5)
'No CTPD computation for over 44hrs.

'Original Logic
'Forfeit OT Hours
'Create CTEX - Overtime Expired code if not already existing
'1. Retrieve OT records from HR_ATTENDANCE where AD_BANKHRS_EXP < Pay Period End Date
'2. For each record retrieved,
'   - remove AD_BANKHRS_EXP value, Frozen (AD_INDICATOR (Incentive)), LDate(NOW), LUser, LTime,
'   - Create a CTEX record for the same hours for the Pay Period End Date, Frozen(AD_INDICATOR (Incentive)).
    
'OT Hours Adjustments
'Create CTRV - Reversed for weekly adjustment code if not already existing
'Create CTPD - Overtime Paid code if not already existing
'1. Retrieve OT Records from HR_ATTENDANCE where AD_BANKHRS_EXP is not blank or null, and
'   Not Frozen - for the Pay Period just marked as closed.
'2. Calculate the Total Hours Worked for the week, i.e. OT hours retrieved + Hours/Week (35)
'3. Create Reversal Entry (CTRV) for total OT retrieved for the week as of Pay End Date
'4. For each record retrieved, make the following adjustment:
'   - >35 < 40  (4hrs - OT * 1)
'   - >=40 < 44 (5hrs - OT15 * 1.5)
'   - >=44      (CTPD * 1.5)
'  If Total Hours Worked - 35 >= 4 then OT Hours * 1
'       - Create an OT Record for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'  If Total Hours Worked - 35 > 4 and <= 8 then 4 OT Hours * 1, (OT Hours - 4) * 1.5
'       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'       - Create an OT15 Record for (OT Hours - 4) * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'  If Total Hours Worked - 35 > 8 then 4 OT Hours * 1, 4 OT Hours * 1.5, (OT Hours - 8) * 1.5
'       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'       - Create an OT15 Record for 4 hrs * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'       - Create an OTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen, Null Expiry Date, Consumed Fully
'           - this is an Overtime Paid adjustment entry for which an automatic CTPD will be done.
'       - Create an CTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen
'           - for OTPD - Overtime Paid Out. This will balance out the OT Outstanding otherwise there will
'             always be an overused CT balance.
'  Update the original OT record retrieved
'       - Remove the Expiry Date
'       - Freeze it

'Adjusted OT record Consumed by CT
'   - Retrieve CT records which are not frozen and not fully adjusted.
'   - Retrieve OT records which are Frozen, Not fully consumed and with Expiry Date
'       - Update AD_CONSUMED with CT hours consumed from this OT record
'           - If fully consumed, remove Expiry Date and Add Comments as used OT
'       - Move through each OT record until CT is fully adjusted
'           - Update AD_Consumed with hours adjusted

'Ticket #21411 - OTRE - Overtime Reinstated. These are expired hours that the user will be re-entering
'back so the employees can take it. They don't want employees to loose out hours because they are still
'trying to understand the logic of OT expired and also they are not really aware of how many hours they
'have.

    Dim SQLQ As String
    Dim rsAttend As New ADODB.Recordset
    Dim rsAddAttend As New ADODB.Recordset
    Dim rsConsAttend As New ADODB.Recordset
    Dim OTHours As Double
    Dim HrsWorked As Double
    Dim StdHrsWeek As Double
    Dim OT15Hrs As Double
    Dim CTPDHrs As Double
    Dim AdjHrs As Double
    Dim OTBal As Double
    Dim CTBal As Double
    Dim xExpiryDate
    
    
    OTHours = 0
    AdjHrs = 0
    HrsWorked = 0
    StdHrsWeek = 35
    'Ticket #23275 - OT Logic Revised.
    'OT15Hrs = 40
    OT15Hrs = 37.5
    CTPDHrs = 44
    
    'Ticket #23275 - They asked us to remove this logic
'    'Forfeit Hours Begin --------------------------------------------------------------------------------
'    'Create CTEX - Overtime Expired code if not already existing
'    Call CreateTableMasterCode("ADRE", "CTEX", "Overtime Expired")
'
'    'Retrieve OT Attendance records which have expired and Frozen
'    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
'    SQLQ = SQLQ & " AND AD_BANKHRS_EXP <" & Date_SQL(xPayEndDate)
'    SQLQ = SQLQ & " AND AD_REASON LIKE ('OT%')"
'    SQLQ = SQLQ & " AND AD_INDICATOR = 1"
'    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_BANKHRS_EXP, AD_DOA"
'    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    Do While Not rsAttend.EOF
'        'Remove the expiry date from these records
'        xExpiryDate = rsAttend("AD_BANKHRS_EXP")
'        rsAttend("AD_BANKHRS_EXP") = Null
'        'rsAttend("AD_INDICATOR") = 1  'Freezing record
'        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & Round((rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))), 2) & " Overtime Hours Expired and " & IIf(IsNull(rsAttend("AD_CONSUMED")), 0, Round(rsAttend("AD_CONSUMED"), 2)) & " Overtime Hours Consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAttend("AD_LDATE") = Date
'        rsAttend("AD_LTIME") = Time$
'        rsAttend("AD_LUSER") = glbUserID
'        rsAttend.Update
'
'        'Create a CTEX record for the same hours
'        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
'        rsAddAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'        rsAddAttend.AddNew
'        rsAddAttend("AD_COMPNO") = "001"
'        rsAddAttend("AD_EMPNBR") = rsAttend("AD_EMPNBR")
'        rsAddAttend("AD_SALARY") = rsAttend("AD_SALARY")
'        rsAddAttend("AD_SALCD") = rsAttend("AD_SALCD")
'        rsAddAttend("AD_JOB") = rsAttend("AD_JOB")
'        rsAddAttend("AD_DHRS") = rsAttend("AD_DHRS")
'        rsAddAttend("AD_WHRS") = rsAttend("AD_WHRS")
'        rsAddAttend("AD_SUPER") = rsAttend("AD_SUPER")
'        rsAddAttend("AD_PAYROLL_ID") = rsAttend("AD_PAYROLL_ID")
'        rsAddAttend("AD_SHIFT") = rsAttend("AD_SHIFT")
'        rsAddAttend("AD_GLNO") = rsAttend("AD_GLNO")
'        rsAddAttend("AD_ORG") = rsAttend("AD_ORG")
'        rsAddAttend("AD_DOA") = xPayEndDate
'        rsAddAttend("AD_REASON") = "CTEX"
'        rsAddAttend("AD_HRS") = rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))
'        rsAddAttend("AD_CONSUMED") = rsAddAttend("AD_HRS")
'        rsAddAttend("AD_COMM") = "Weekly Forfeited Hours for the Attendance record '" & rsAttend("AD_REASON") & "' dated '" & Format(rsAttend("AD_DOA"), "Short Date") & "' with Expiry Date of '" & Format(xExpiryDate, "Short Date") & "." & vbCrLf & Round((rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))), 2) & " Overtime Hour(s) Expired and " & IIf(IsNull(rsAttend("AD_CONSUMED")), 0, Round(rsAttend("AD_CONSUMED"), 2)) & " Overtime Hour(s) Consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
'        rsAddAttend("AD_INDICATOR") = 1  'Freezing record
'        rsAddAttend("AD_LUSER") = glbUserID
'        rsAddAttend("AD_LDATE") = Date
'        rsAddAttend("AD_LTIME") = Time$
'        rsAddAttend("AD_SOURCE") = "IHRFOR"
'        rsAddAttend.Update
'        rsAddAttend.Close
'        Set rsAddAttend = Nothing
'
'        rsAttend.MoveNext
'    Loop
'    rsAttend.Close
'    Set rsAttend = Nothing
'    'Forfeit Hours End ----------------------------------------------------------------------------------
    
    'OT Hours Adjustments Begin -------------------------------------------------------------------------
    'Create CTRV - OT Reversed code, CTPD - Overtime Paid code, if not already existing
    Call CreateTableMasterCode("ADRE", "CTRV", "Overtime Reversed")
    Call CreateTableMasterCode("ADRE", "OTPD", "Overtime Earned for Pay Out")
    Call CreateTableMasterCode("ADRE", "CTPD", "Overtime Paid Out")
    Call CreateTableMasterCode("ADRE", "OT15", "Overtime at Time and One-Half")
    'Ticket #24201 - Put the Adjusted hours in the different OT code (OT1) to make report easy to read.
    Call CreateTableMasterCode("ADRE", "OT1", "Overtime at Straight Time")
    
    'Ticket #21197 Begin ------------------------------------------------------------------------
    'Adjust/Consume the OTs and CTs of the Pay Week before Reversing the original entry
    'Retrieve CT Attendance records for the Pay Period just closed where Not Frozen
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE "
    SQLQ = SQLQ & " AD_DOA >=" & Date_SQL(xPayStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
    SQLQ = SQLQ & " AND AD_REASON = 'CT'"
    SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR ASC, AD_DOA ASC"
    rsConsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    Do While Not rsConsAttend.EOF
        
        'Retrieve OT Attendance for the Pay Period not Consumed fully, not expired
        'and not Frozen
        'Ticket #24035 - Fixing the logic as Expiry Date (AD_BANKHRS_EXP) column update in ESS has been disabled but
        'the logic in this routine dependent on having a date on this field. It does not matter what date.
        'SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (AD_BANKHRS_EXP IS NOT NULL OR AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
        SQLQ = SQLQ & " AND AD_EMPNBR = " & rsConsAttend("AD_EMPNBR")
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
        SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
        SQLQ = SQLQ & " AND AD_REASON LIKE 'OT%'"
        'Ticket #24035 - added above instead
        'SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
        SQLQ = SQLQ & " ORDER BY AD_DOA ASC"
        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Do While Not rsAttend.EOF
            'If fully consumed - Remove Expiry Date, add Comments
            
            'Check if there is enough hours OT (Ad_hrs  - Consumed) left to be
            'consumed by CT (Ad_Hrs - Consumed) left
            OTBal = 0
            CTBal = 0
            OTBal = (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")))
            CTBal = rsConsAttend("AD_HRS") - IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED"))
            If OTBal >= CTBal Then
                'Enough OT Hours left to consume CT Bal
                'Update OT Consumed with all of CT Bal - some OT will still be left
                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + CTBal
                
                'Update CT Consumed with CT Bal - nothing will be left
                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + CTBal
                
                'Update other fields of both these record sets
                'OT recordset
                If OTBal = CTBal Then
                    'Since all OT Hours has been used up, remove expiry date and update Comments as consumed
                    rsAttend("AD_BANKHRS_EXP") = Null
                    rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All of original OT hours consumed within the week of " & Format(xPayStartDate, "mmm dd, yyyy") & " - " & Format(xPayEndDate, "mmm dd, yyyy") & "."
                    rsAttend("AD_INDICATOR") = 1  'Freezing record
                End If
                rsAttend("AD_LDATE") = Date
                rsAttend("AD_LTIME") = Time$
                rsAttend("AD_LUSER") = glbUserID
                rsAttend.Update
                
                'CT recordset
                rsConsAttend("AD_INDICATOR") = 1  'Freezing record
                rsConsAttend("AD_LDATE") = Date
                rsConsAttend("AD_LTIME") = Time$
                rsConsAttend("AD_LUSER") = glbUserID
                rsConsAttend.Update
                
                'All CT hours have balanced up with OT - Move to next CT record
                Exit Do
            Else
                'OT Hours left to consume is less that CT Bal to consume
                'Update CT Consume with OT hours left to be consumed (OTBal)
                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + OTBal
                
                'Update OT Consume with OT hours left to be consumed (OTBal) = all OT hours used up
                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + OTBal
                
                'Update OT recordset
                'Since all OT hours has been used up, remove expiry date, update Comments as consumed
                rsAttend("AD_BANKHRS_EXP") = Null
                rsAttend("AD_INDICATOR") = 1  'Freezing record
                rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All of original OT hours consumed within the week of " & Format(xPayStartDate, "mmm dd, yyyy") & " - " & Format(xPayEndDate, "mmm dd, yyyy") & "."
                rsAttend("AD_LDATE") = Date
                rsAttend("AD_LTIME") = Time$
                rsAttend("AD_LUSER") = glbUserID
                rsAttend.Update
            
                'Update CT recordset
                'rsConsAttend("AD_INDICATOR") = 1  'Freezing record
                rsConsAttend("AD_LDATE") = Date
                rsConsAttend("AD_LTIME") = Time$
                rsConsAttend("AD_LUSER") = glbUserID
                rsConsAttend.Update
                                
            End If
            
            rsAttend.MoveNext
        Loop
        rsAttend.Close
        Set rsAttend = Nothing
    
        rsConsAttend.MoveNext
    Loop
    rsConsAttend.Close
    Set rsConsAttend = Nothing
    'Ticket #21197 End --------------------------------------------------------------------------

    
    'OT Reversal Entry
    'Retrieve OT Attendance records for the Pay Period just closed
    'where AD_BANKHRS_EXP is not blank or null, and Not Frozen
    SQLQ = "SELECT SUM(CASE WHEN AD_HRS IS NULL THEN 0 ELSE AD_HRS END) - SUM(CASE WHEN AD_CONSUMED IS NULL THEN 0 ELSE AD_CONSUMED END) AS TOT_OT, AD_EMPNBR "
    'Ticket #24035 - Fixing the logic as Expiry Date (AD_BANKHRS_EXP) column update in ESS has been disabled but
    'the logic in this routine dependent on having a date on this field. It does not matter what date.
    'SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE (AD_BANKHRS_EXP IS NOT NULL OR AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
    SQLQ = SQLQ & " AND AD_REASON = 'OT'"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Do While Not rsAttend.EOF
        'Calculate the Total Hours Worked for the week, i.e. OT hours retrieved + Hours/Week (35)
        OTHours = IIf(IsNull(rsAttend("TOT_OT")), 0, rsAttend("TOT_OT"))
        HrsWorked = StdHrsWeek + OTHours
        
        'Create CTRV - OT Reversal Entry
        Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "CTRV", OTHours, "", "OT Reversal entry for the adjustment.", OTHours)
        
        
        'Ticket #23275 - OT Logic Revised.
        'New Logic:
        '>35 < 37.5 (2.5hrs OT * 1)
        '>=37.5     (OT15 * 1.5)
        'No CTPD computation for over 44hrs.
        
        'Create the OT Adjustment entries
        '   - >35 < 40  (4hrs - OT * 1)
        '   - >=40 < 44 (5hrs - OT15 * 1.5)
        '   - >=44      (CTPD * 1.5)
        '  If Total Hours Worked - 35 >= 4 then OT Hours * 1
        '       - Create an OT Record for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
        '  If Total Hours Worked - 35 > 4 and <= 8 then 4 OT Hours * 1, (OT Hours - 4) * 1.5
        '       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
        '       - Create an OT15 Record for (OT Hours - 4) * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
        '  If Total Hours Worked - 35 > 8 then 4 OT Hours * 1, 4 OT Hours * 1.5, (OT Hours - 8) * 1.5
        '       - Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
        '       - Create an OT15 Record for 4 hrs * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
        '       - Create an CTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen
        
        'Ticket #23275 - OT Logic Revised >35 < 37.5.
        'If (HrsWorked - StdHrsWeek) <= 4 Then
        If (HrsWorked - StdHrsWeek) <= 2.5 Then
        'If (HrsWorked > StdHrsWeek) And (HrsWorked < OT15Hrs) Then     '> 35 < 40
            'Ajusted OT Hours
            AdjHrs = HrsWorked - StdHrsWeek
            
            'Create an OT Record for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
            If AdjHrs > 0 Then
                'Ticket #24201 - Put the Adjusted hours in the different OT code (OT1) to make report easy to read.
                'Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
                Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT1", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
            End If
        End If
        
        'Ticket #23275 - OT Logic Revised >=37.5
        'If (HrsWorked - StdHrsWeek) > 4 And (HrsWorked - StdHrsWeek) <= 8 Then
        If (HrsWorked - StdHrsWeek) > 2.5 Then 'And (HrsWorked - StdHrsWeek) <= 8 Then
        'If (HrsWorked >= OT15Hrs) And (HrsWorked < CTPDHrs) Then    '>=40 < 44
            'Ajusted OT Hours for OT
            'Ticket #23275 - OT Logic Revised
            'AdjHrs = 4
            AdjHrs = 2.5
            
            'Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
            'Create an OT Record for 2.5 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
            'Ticket #24201 - Put the Adjusted hours in the different OT code (OT1) to make report easy to read.
            'Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT1", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
            
            'Ajusted OT Hours for OT15
            'Ticket #23275 - OT Logic Revised
            'AdjHrs = (HrsWorked - (OT15Hrs - 1)) * 1.5
            AdjHrs = (HrsWorked - OT15Hrs) * 1.5
            
            'Create an OT15 Record for (OT Hours - 4) * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT15", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
        End If
        
        'Ticket #23275 - OT Logic Revised - No CTPD calculation
'        If (HrsWorked - StdHrsWeek) > 8 Then
'        'If (HrsWorked >= CTPDHrs) Then    '>=44
'            'Ajusted OT Hours for OT
'            AdjHrs = 4
'
'            'Create an OT Record for 4 hrs for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'
'
'            'Ajusted OT Hours for OT15
'            AdjHrs = 4 * 1.5
'
'            'Create an OT15 Record for 4 hrs * 1.5 for Pay Period End Date, Frozen, Expiry Date (Pay Period End Date + 30)
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OT15", AdjHrs, DateAdd("d", 30, xPayEndDate), "Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."))
'
'
'            'Ajusted OT Hours for CTPD
'            AdjHrs = (HrsWorked - (CTPDHrs - 1)) * 1.5
'
'            'Create an OTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen, Null Expiry Date, Fully Consumed
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "OTPD", AdjHrs, "", "Overtime Earned for Pay Out Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."), AdjHrs)
'
'            'Create an CTPD Record for (OT Hours - 8) * 1.5 for Pay Period End Date, Frozen
'            Call Add_OT_Adjustment_Attendance(rsAttend("AD_EMPNBR"), xPayEndDate, "CTPD", AdjHrs, "", "Paid Out Adjustment entry for the total of " & Round(OTHours, 2) & IIf(OTHours = 1, " Overtime hour.", " Overtime hours."), AdjHrs)
'        End If
        
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
        
    'Freeze the original OT entries for the Pay Period just closed
    'Retrieve OT Attendance records for the Pay Period just closed
    'where AD_BANKHRS_EXP is not blank or null, and Not Frozen
    'Ticket #24035 - Fixing the logic as Expiry Date (AD_BANKHRS_EXP) column update in ESS has been disabled but
    'the logic in this routine dependent on having a date on this field. It does not matter what date.
    'SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (AD_BANKHRS_EXP IS NOT NULL OR AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
    SQLQ = SQLQ & " AND AD_REASON = 'OT'"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttend.EOF
        'Remove the expiry date from these records and freeze it
        rsAttend("AD_BANKHRS_EXP") = Null
        rsAttend("AD_INDICATOR") = 1  'Freezing record
        rsAttend("AD_CONSUMED") = rsAttend("AD_HRS")
        'rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "Original OT entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))) & " hour(s) from the original " & rsAttend("AD_HRS") & " OT hours entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
        rsAttend("AD_LDATE") = Date
        rsAttend("AD_LTIME") = Time$
        rsAttend("AD_LUSER") = glbUserID
        rsAttend.Update
    
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
    
    'Ticket #21411 - Since they are giving back the expired hours - need to freeze it so it can be consumed.
    'Freeze the OTRE entries for the Pay Period just closed
    'Retrieve OTRE Attendance records for the Pay Period just closed
    'where AD_BANKHRS_EXP is not blank or null, and Not Frozen
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
    SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xPayStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
    SQLQ = SQLQ & " AND AD_REASON = 'OTRE'"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttend.EOF
        'Remove the expiry date from these records and freeze it
        'rsAttend("AD_BANKHRS_EXP") = Null
        rsAttend("AD_INDICATOR") = 1  'Freezing record
        'rsAttend("AD_CONSUMED") = rsAttend("AD_HRS")
        'rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "Original OT entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & rsAttend("AD_HRS") & " OTRE hours entry frozen as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
        rsAttend("AD_LDATE") = Date
        rsAttend("AD_LTIME") = Time$
        rsAttend("AD_LUSER") = glbUserID
        rsAttend.Update
    
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
    
    'Freeze the CT entries for the Pay Period just closed
    'Retrieve CT Attendance records for the Pay Period just closed
    'where Not Frozen
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE "
    SQLQ = SQLQ & " AD_DOA >=" & Date_SQL(xPayStartDate)
    SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
    SQLQ = SQLQ & " AND AD_INDICATOR <> 1"
    SQLQ = SQLQ & " AND AD_REASON = 'CT'"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttend.EOF
        'Freeze it
        'rsAttend("AD_BANKHRS_EXP") = Null
        rsAttend("AD_INDICATOR") = 1  'Freezing record
        'rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "Original CT entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
        rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED"))) & " hour(s) from the original " & rsAttend("AD_HRS") & " CT hours entry frozen for the adjustment as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
        rsAttend("AD_LDATE") = Date
        rsAttend("AD_LTIME") = Time$
        rsAttend("AD_LUSER") = glbUserID
        rsAttend.Update
    
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
    
    'OT Hours Adjustments End ---------------------------------------------------------------------------

    'Consumed Hours (Adjust CT against OTs as Consumed Hours) Begin --------------------------------------
    'Retrieve CT Attendance records, not fully adjusted
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE "
    'SQLQ = SQLQ & " AD_DOA >=" & Date_SQL(xPayStartDate)
    SQLQ = SQLQ & " AD_DOA <=" & Date_SQL(xPayEndDate)
    SQLQ = SQLQ & " AND AD_INDICATOR = 1"
    SQLQ = SQLQ & " AND AD_REASON = 'CT'"
    SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR ASC, AD_DOA ASC"
    rsConsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    Do While Not rsConsAttend.EOF
        
        'Retrieve OT Attendance not Consumed fully, not expired and Frozen
        'Ticket #24035 - Fixing the logic as Expiry Date (AD_BANKHRS_EXP) column update in ESS has been disabled but
        'the logic in this routine dependent on having a date on this field. It does not matter what date.
        'SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_BANKHRS_EXP IS NOT NULL "
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (AD_BANKHRS_EXP IS NOT NULL OR AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
        SQLQ = SQLQ & " AND AD_EMPNBR = " & rsConsAttend("AD_EMPNBR")
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xPayEndDate)
        SQLQ = SQLQ & " AND AD_INDICATOR = 1"
        SQLQ = SQLQ & " AND AD_REASON LIKE 'OT%'"
        'Ticket #24035 - added above instead
        'SQLQ = SQLQ & " AND (AD_HRS <> AD_CONSUMED OR AD_CONSUMED IS NULL)"
        SQLQ = SQLQ & " ORDER BY AD_DOA ASC"
        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Do While Not rsAttend.EOF
            'If fully consumed - Remove Expiry Date, add Comments
            
            'Check if there is enough hours OT (Ad_hrs  - Consumed) left to be consumed by CT (Ad_Hrs - Consumed) left
            OTBal = 0
            CTBal = 0
            OTBal = (rsAttend("AD_HRS") - IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")))
            CTBal = rsConsAttend("AD_HRS") - IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED"))
            If OTBal >= CTBal Then
                'Enough OT Hours left to consume CT Bal
                'Update OT Consumed with all of CT Bal - some OT will still be left
                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + CTBal
                
                'Update CT Consumed with CT Bal - nothing will be left
                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + CTBal
                
                'Update other fields of both these record sets
                'OT recordset
                If OTBal = CTBal Then
                    'Since all OT Hours has been used up, remove expiry date and update Comments as consumed
                    rsAttend("AD_BANKHRS_EXP") = Null
                    rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All OT hours consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
                End If
                rsAttend("AD_LDATE") = Date
                rsAttend("AD_LTIME") = Time$
                rsAttend("AD_LUSER") = glbUserID
                rsAttend.Update
                
                'CT recordset
                rsConsAttend("AD_INDICATOR") = 1  'Freezing record
                rsConsAttend("AD_LDATE") = Date
                rsConsAttend("AD_LTIME") = Time$
                rsConsAttend("AD_LUSER") = glbUserID
                rsConsAttend.Update
                
                'All CT hours have balanced up with OT - Move to next CT record
                Exit Do
            Else
                'OT Hours left to consume is less that CT Bal to consume
                'Update CT Consume with OT hours left to be consumed (OTBal)
                rsConsAttend("AD_CONSUMED") = IIf(IsNull(rsConsAttend("AD_CONSUMED")), 0, rsConsAttend("AD_CONSUMED")) + OTBal
                
                'Update OT Consume with OT hours left to be consumed (OTBal) = all OT hours used up
                rsAttend("AD_CONSUMED") = IIf(IsNull(rsAttend("AD_CONSUMED")), 0, rsAttend("AD_CONSUMED")) + OTBal
                
                'Update OT recordset
                'Since all OT hours has been used up, remove expiry date, update Comments as consumed
                rsAttend("AD_BANKHRS_EXP") = Null
                rsAttend("AD_COMM") = IIf(IsNull(rsAttend("AD_COMM")), "", rsAttend("AD_COMM") & vbCrLf & "") & "All OT hours consumed as of " & Format(xPayEndDate, "mmm dd, yyyy") & "."
                rsAttend("AD_LDATE") = Date
                rsAttend("AD_LTIME") = Time$
                rsAttend("AD_LUSER") = glbUserID
                rsAttend.Update
            
                'Update CT recordset
                rsConsAttend("AD_INDICATOR") = 1  'Freezing record
                rsConsAttend("AD_LDATE") = Date
                rsConsAttend("AD_LTIME") = Time$
                rsConsAttend("AD_LUSER") = glbUserID
                rsConsAttend.Update
                                
            End If
            
            rsAttend.MoveNext
        Loop
        rsAttend.Close
        Set rsAttend = Nothing
    
        rsConsAttend.MoveNext
    Loop
    rsConsAttend.Close
    Set rsConsAttend = Nothing
    'Consumed Hours (Adjust CT against OTs as Consumed Hours) End ----------------------------------------

    
End Sub

Private Sub Add_OT_Adjustment_Attendance(xEmpnbr, xDate, xReason, xAdjHrs, xExpiryDate, xComments, Optional xConsumed)
    Dim rsAddAttend As New ADODB.Recordset
    Dim rsCurSal As New ADODB.Recordset
    Dim rsCurPos As New ADODB.Recordset
    Dim SQLQ As String
    
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE 1 = 2"
        rsAddAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsAddAttend.AddNew
        rsAddAttend("AD_COMPNO") = "001"
        rsAddAttend("AD_EMPNBR") = xEmpnbr
                
        'Update with Salary info.
        SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEmpnbr
        rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCurSal.BOF Then
            rsAddAttend("AD_SALARY") = rsCurSal("SH_SALARY")
            rsAddAttend("AD_SALCD") = rsCurSal("SH_SALCD")
        End If
        rsCurSal.Close
        Set rsCurSal = Nothing
        
        'Update with Position info.
        SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS,JH_REPTAU,JH_PAYROLL_ID,JH_SHIFT,JH_GLNO,JH_ORG FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEmpnbr
        rsCurPos.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsCurPos.EOF Then
            rsAddAttend("AD_JOB") = rsCurPos("JH_JOB")
            rsAddAttend("AD_DHRS") = rsCurPos("JH_DHRS")
            rsAddAttend("AD_WHRS") = rsCurPos("JH_WHRS")
            rsAddAttend("AD_SUPER") = rsCurPos("JH_REPTAU")
            rsAddAttend("AD_PAYROLL_ID") = rsCurPos("JH_PAYROLL_ID")
            rsAddAttend("AD_SHIFT") = rsCurPos("JH_SHIFT")
            rsAddAttend("AD_GLNO") = rsCurPos("JH_GLNO")
            rsAddAttend("AD_ORG") = rsCurPos("JH_ORG")
        End If
        rsCurPos.Close
        Set rsCurPos = Nothing
        
        rsAddAttend("AD_DOA") = xDate
        rsAddAttend("AD_REASON") = xReason
        rsAddAttend("AD_HRS") = xAdjHrs
        rsAddAttend("AD_COMM") = xComments
        rsAddAttend("AD_BANKHRS_EXP") = IIf(IsDate(xExpiryDate), xExpiryDate, Null)
        rsAddAttend("AD_INDICATOR") = 1  'Freezing record
        rsAddAttend("AD_CONSUMED") = IIf(IsMissing(xConsumed), Null, xConsumed)
        rsAddAttend("AD_LUSER") = glbUserID
        rsAddAttend("AD_LDATE") = Date
        rsAddAttend("AD_LTIME") = Time$
        rsAddAttend("AD_SOURCE") = "IHRADJ"
        rsAddAttend.Update
        rsAddAttend.Close
        Set rsAddAttend = Nothing

End Sub

Private Function AllowSaveChanges() As Boolean
    Dim SQLQ As String
    Dim rsPayPeriod As New ADODB.Recordset
    Dim xStatus As String
    Dim X%
    
    'Initialise to allow.
    AllowSaveChanges = True
    xPPNoChanged = ""
    
    'Get the WHERE clause of the existing Pay Period record being updated.
    Call getWSQLQ("O")
    
    'Check if the Pay Period record exists and then compare the Pay Period dates of the saved record with the
    'on screen data to find out if any of the dates changed.
    For X% = 0 To txtSeq.count - 1
        If Len(txtSeq(X%)) > 0 Then
            SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE " & fglbVSQLQ
            SQLQ = SQLQ & " AND PP_NBR = " & txtSeq(X)
            rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPayPeriod.EOF Then
                'rsVE("PP_NBR") = txtSeq(X)
                'rsVE("PP_START") = dlpFrom(X%)
                'rsVE("PP_END") = dlpTo(X%)
                
                'Check if Pay Period has changed.
                If rsPayPeriod("PP_START") <> dlpFrom(X%) Or rsPayPeriod("PP_END") <> dlpTo(X%) Then
                    'This Pay Period has changed, check if any Timesheet records exists for saved Pay Period
                    xStatus = get_TSStatus(rsPayPeriod("PP_START"), rsPayPeriod("PP_END"), rsPayPeriod("PP_NBR"))
                    
                    If xStatus <> "Not Entered" Then
                        'There are Timesheet records for the original Pay Period. Do not allow to save the changes.
                        AllowSaveChanges = False
                        
                        xPPNoChanged = txtSeq(X)
                        
                        Exit For
                    End If
                Else
                    'This Pay Period has not changed, check the next one
                End If
            End If
            rsPayPeriod.Close
            Set rsPayPeriod = Nothing
        End If
    Next
    
End Function

Private Function get_TSStatus(strPPStartDate, strPPEndDate, xPPNbr)
Dim SQLQ, statusFlag
Dim rsDS As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim gdbESS As New ADODB.Connection
Dim xEmpStatus As String
Dim xLOA As Boolean

    If glbSQL Or glbOracle Then
        Set gdbESS = gdbAdoIhr001
    Else
        gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
    End If
    
    On Error Resume Next
    SQLQ = "SELECT DISTINCT AD_APPROVED,AD_UPLOAD FROM HR_TIMESHEET "
    'SQLQ = SQLQ & " WHERE AD_EMPNBR =" & xEmpnbr
    SQLQ = SQLQ & " WHERE AD_PPSTART >=" & Date_SQL(strPPStartDate)
    SQLQ = SQLQ & " AND AD_PPEND <=" & Date_SQL(strPPEndDate)
    SQLQ = SQLQ & " AND AD_PPNBR =" & xPPNbr

    If glbSQL Or glbOracle Then
        rsDS.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        rsDS.Open SQLQ, gdbESS, adOpenForwardOnly
    End If
    
    get_TSStatus = ""
    statusFlag = True
    
    If rsDS.EOF Then
        get_TSStatus = "Not Entered"
    Else
        Do While Not rsDS.EOF
            If statusFlag Then
                If IsNull(rsDS("AD_APPROVED")) Then
                    If rsDS("AD_UPLOAD") & "" = "Y" Then
                        get_TSStatus = "SUBMITTED"
                    Else
                        get_TSStatus = "SAVED"
                    End If
                Else
                    get_TSStatus = rsDS("AD_APPROVED")
                    
                    If get_TSStatus = "RESUBMIT" Then get_TSStatus = "RESUBMITTED"
                End If
                statusFlag = False
                Exit Do
            Else
                get_TSStatus = "SAVED"
                Exit Do
            End If
            rsDS.MoveNext
        Loop
    End If
    rsDS.Close
    Set rsDS = Nothing
    
End Function

Private Function Closing_PayPeriod() As Boolean
    Dim SQLQ As String
    Dim rsPayPeriod As New ADODB.Recordset
    Dim xClosed As Boolean
    Dim X%
    
    'Initialise to allow.
    Closing_PayPeriod = False
    
    xPPNoChanged = ""
    
    'Get the WHERE clause of the existing Pay Period record being updated.
    Call getWSQLQ("O")
    
    'Check if the Pay Period record exists and then compare the Pay Period dates of the saved record with the
    'on screen data to find out if any of the dates changed.
    For X% = 0 To txtSeq.count - 1
        If Len(txtSeq(X%)) > 0 Then
            SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE " & fglbVSQLQ
            SQLQ = SQLQ & " AND PP_NBR = " & txtSeq(X)
            rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsPayPeriod.EOF Then
                'rsVE("PP_NBR") = txtSeq(X)
                'rsVE("PP_START") = dlpFrom(X%)
                'rsVE("PP_END") = dlpTo(X%)
                
                'Check if closing Pay Period
                If chkUploaded(X%) Then
                    xClosed = True
                Else
                    xClosed = False
                End If
                If rsPayPeriod("PP_UPLOADED") <> xClosed And xClosed = True Then
                    Closing_PayPeriod = True
                    Exit For
                Else
                    'This Pay Period has not changed, check the next one
                End If
            End If
            rsPayPeriod.Close
            Set rsPayPeriod = Nothing
        End If
    Next
    
End Function

