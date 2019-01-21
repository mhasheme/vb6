VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRESSTrnAudit 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "ESS Requests - Transaction Audit Report"
   ClientHeight    =   10485
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   11910
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
   ScaleHeight     =   10485
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   54
      Top             =   9915
      Width           =   9855
   End
   Begin VB.VScrollBar scrControl 
      Height          =   9315
      LargeChange     =   315
      Left            =   9960
      Max             =   100
      SmallChange     =   315
      TabIndex        =   53
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9360
      Top             =   9120
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
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   9735
      Begin VB.ComboBox comGroup 
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
         Index           =   0
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "First Level of grouping records"
         Top             =   8325
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
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
         Index           =   1
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Tag             =   "Second level of grouping records"
         Top             =   8662
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   10
         Tag             =   "00-Employee Position Shift"
         Top             =   3820
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox comGroup 
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
         Index           =   2
         ItemData        =   "FZESSTrnAudit.frx":0000
         Left            =   1890
         List            =   "FZESSTrnAudit.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Tag             =   "Second level of grouping records"
         Top             =   9000
         Width           =   2325
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   390
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "n/a"
         MaxLength       =   0
         LookupType      =   1
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   1920
         TabIndex        =   9
         Tag             =   "00-Enter Section Code"
         Top             =   3477
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   8
         Tag             =   "00-Enter Administered By Code"
         Top             =   3134
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
         MaxLength       =   10
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   7
         Tag             =   "00-Enter Region Code"
         Top             =   2791
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   2105
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDPT"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "00-Enter Status Code"
         Top             =   1762
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1419
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDOR"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   1076
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   733
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "n/a"
         MaxLength       =   0
         LookupType      =   2
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkShowVacReq 
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Tag             =   "If X-Show Vacation Requests"
         Top             =   5340
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Vacation Requests  "
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
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Tag             =   "10-Enter Employee Number"
         Top             =   2448
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkShowTimeOffReq 
         Height          =   225
         Left            =   360
         TabIndex        =   16
         Tag             =   "If X-Show Time Off Requests"
         Top             =   5640
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Time Off Requests     "
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
      Begin INFOHR_Controls.DateLookup dlpReqTo 
         Height          =   285
         Left            =   3960
         TabIndex        =   12
         Tag             =   "40-Date upto and including this date forward"
         Top             =   4163
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpReqFrom 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Tag             =   "40-Date from and including this date forward"
         Top             =   4163
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpTo 
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Tag             =   "40-Date upto and including this date forward"
         Top             =   4515
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpFrom 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Tag             =   "40-Date from and including this date forward"
         Top             =   4515
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSCheck chkShowApproved 
         Height          =   225
         Left            =   360
         TabIndex        =   17
         Tag             =   "If X-Show Vacation Requests"
         Top             =   6480
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Approved"
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
      Begin Threed.SSCheck chkShowArchived 
         Height          =   225
         Left            =   360
         TabIndex        =   18
         Tag             =   "If X-Show Vacation Requests"
         Top             =   6840
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Archived"
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
      Begin Threed.SSCheck chkShowDeleted 
         Height          =   225
         Left            =   360
         TabIndex        =   19
         Tag             =   "If X-Show Vacation Requests"
         Top             =   7200
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Deleted"
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
      Begin Threed.SSCheck chkShowDelApproved 
         Height          =   225
         Left            =   360
         TabIndex        =   20
         Tag             =   "If X-Show Vacation Requests"
         Top             =   7560
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Delete Approved"
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
      Begin Threed.SSCheck chkShowModified 
         Height          =   225
         Left            =   3360
         TabIndex        =   21
         Tag             =   "If X-Show Vacation Requests"
         Top             =   6480
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Modified"
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
      Begin Threed.SSCheck chkShowNewRequest 
         Height          =   225
         Left            =   3360
         TabIndex        =   22
         Tag             =   "If X-Show Vacation Requests"
         Top             =   6840
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show New Request"
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
      Begin Threed.SSCheck chkShowRejected 
         Height          =   225
         Left            =   3360
         TabIndex        =   23
         Tag             =   "If X-Show Vacation Requests"
         Top             =   7200
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Rejected"
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
      Begin Threed.SSCheck chkShowResubmitted 
         Height          =   225
         Left            =   3360
         TabIndex        =   24
         Tag             =   "If X-Show Vacation Requests"
         Top             =   7560
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Resubmitted"
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
      Begin Threed.SSCheck chkShowAppForward 
         Height          =   225
         Left            =   6000
         TabIndex        =   25
         Tag             =   "If X-Show Vacation Requests"
         Top             =   6480
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Approved Forward"
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
      Begin Threed.SSCheck chkShowAppFailed 
         Height          =   225
         Left            =   6000
         TabIndex        =   26
         Tag             =   "If X-Show Vacation Requests"
         Top             =   6840
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "  Show Approved Failed"
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
         Left            =   120
         TabIndex        =   52
         Top             =   435
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
         Left            =   120
         TabIndex        =   51
         Top             =   778
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
         Left            =   120
         TabIndex        =   50
         Top             =   1464
         Width           =   420
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Left            =   120
         TabIndex        =   49
         Top             =   1807
         Width           =   450
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
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
         Left            =   120
         TabIndex        =   48
         Top             =   2493
         Width           =   1290
      End
      Begin VB.Label lblSelCri 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   47
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #1"
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
         Left            =   120
         TabIndex        =   46
         Top             =   8385
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Final Sort"
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
         Left            =   120
         TabIndex        =   45
         Top             =   9060
         Width           =   660
      End
      Begin VB.Label lblRepGrp 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Grouping"
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
         TabIndex        =   44
         Top             =   8115
         Width           =   1575
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
         Left            =   120
         TabIndex        =   43
         Top             =   1121
         Width           =   615
      End
      Begin VB.Label lblAdmin 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Administered By"
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
         Left            =   120
         TabIndex        =   42
         Top             =   3179
         Width           =   1125
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
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
         Left            =   120
         TabIndex        =   41
         Top             =   2836
         Width           =   510
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   40
         Top             =   3522
         Width           =   540
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
         Left            =   120
         TabIndex        =   39
         Top             =   2150
         Width           =   630
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
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
         Left            =   120
         TabIndex        =   38
         Top             =   3865
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction From"
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
         Left            =   120
         TabIndex        =   37
         Top             =   4560
         Width           =   1230
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Request From"
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
         Left            =   120
         TabIndex        =   36
         Top             =   4208
         Width           =   990
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3600
         TabIndex        =   35
         Top             =   4208
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3600
         TabIndex        =   34
         Top             =   4560
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Request Types:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   5040
         Width           =   1125
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction Types:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   6120
         Width           =   1365
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #2"
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
         Left            =   120
         TabIndex        =   31
         Top             =   8722
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmRESSTrnAudit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowEmp_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm("ESS Requests - Transaction Audit Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
      Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
      Call set_PrintState(True)
End If
Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub comEmpType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(1).AddItem "Date/Time of Action"
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0

    comGroup(2).AddItem "Employee Name"
    comGroup(2).AddItem "Employee Number"
    comGroup(2).ListIndex = 0

End Sub

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ORG"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    End Select
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
    End If
End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(getEmpnbr(elpEEID.Text)) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
End If


If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ReqFTDates()
Dim TempCri As String
Dim dtYYY As Integer, dtMM As Integer, dtDD As Integer

If Len(dlpReqFrom.Text) = 0 And Len(dlpReqTo.Text) = 0 Then Exit Sub

If Len(dlpReqFrom.Text) > 0 And Len(dlpReqTo.Text) > 0 Then
    TempCri = "({HR_REQAUDIT.RT_FROM} "
    dtYYY = Year(dlpReqFrom.Text)
    dtMM = month(dlpReqFrom.Text)
    dtDD = Day(dlpReqFrom.Text)
    TempCri = TempCri & " >= Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ") "
    dtYYY = Year(dlpReqTo.Text)
    dtMM = month(dlpReqTo.Text)
    dtDD = Day(dlpReqTo.Text)
    TempCri = TempCri & " AND {HR_REQAUDIT.RT_TO} <= Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
Else
    If Len(dlpReqFrom.Text) > 0 Then
        TempCri = "({HR_REQAUDIT.RT_FROM} "
        TempCri = TempCri & " >= "
        dtYYY = Year(dlpReqFrom.Text)
        dtMM = month(dlpReqFrom.Text)
        dtDD = Day(dlpReqFrom.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
    If Len(dlpReqTo.Text) > 0 Then
        TempCri = "({HR_REQAUDIT.RT_TO} "
        TempCri = TempCri & " <= "
        dtYYY = Year(dlpReqTo.Text)
        dtMM = month(dlpReqTo.Text)
        dtDD = Day(dlpReqTo.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & TempCri
Else
    glbstrSelCri = TempCri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_TransFTDates()
Dim TempCri As String
Dim dtYYY As Integer, dtMM As Integer, dtDD As Integer

If Len(dlpFrom.Text) = 0 And Len(dlpTo.Text) = 0 Then Exit Sub

TempCri = "({HR_REQAUDIT.RT_LDATE} "
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dtYYY = Year(dlpFrom.Text)
    dtMM = month(dlpFrom.Text)
    dtDD = Day(dlpFrom.Text)
    TempCri = TempCri & " in Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ") "
    dtYYY = Year(dlpTo.Text)
    dtMM = month(dlpTo.Text)
    dtDD = Day(dlpTo.Text)
    TempCri = TempCri & " to Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
Else
    If Len(dlpFrom.Text) > 0 Then
        TempCri = TempCri & " >= "
        dtYYY = Year(dlpFrom.Text)
        dtMM = month(dlpFrom.Text)
        dtDD = Day(dlpFrom.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
    If Len(dlpTo.Text) > 0 Then
        TempCri = TempCri & " <= "
        dtYYY = Year(dlpTo.Text)
        dtMM = month(dlpTo.Text)
        dtDD = Day(dlpTo.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & TempCri
Else
    glbstrSelCri = TempCri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_VacTimeOffReq()
Dim EECri As String

'Show Vacation Requests?
If Not chkShowVacReq Then
    EECri = "{HR_REQAUDIT.RT_REQTYPE} <> 'VR'"

    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    
    glbiOneWhere = True
End If


'Show Time Off Requests?
If Not chkShowTimeOffReq Then
    EECri = "{HR_REQAUDIT.RT_REQTYPE} <> 'TR'"

    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
    
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_TransactionTypes()
Dim EECri As String
Dim xTransType As String
Dim xNotTransType As String

'Transactions Types
xTransType = ""
xNotTransType = ""

If chkShowApproved Then
    xTransType = "'AP'"
Else
    xNotTransType = "'AP'"
End If

If chkShowArchived Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'AR'"
    Else
        xTransType = "'AR'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'AR'"
    Else
        xNotTransType = "'AR'"
    End If
End If

If chkShowDeleted Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'DB'"
    Else
        xTransType = "'DB'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'DB'"
    Else
        xNotTransType = "'DB'"
    End If
End If

If chkShowDelApproved Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'DA'"
    Else
        xTransType = "'DA'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'DA'"
    Else
        xNotTransType = "'DA'"
    End If
End If

If chkShowModified Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'MB'"
    Else
        xTransType = "'MB'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'MB'"
    Else
        xNotTransType = "'MB'"
    End If
End If

If chkShowNewRequest Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'NR'"
    Else
        xTransType = "'NR'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'NR'"
    Else
        xNotTransType = "'NR'"
    End If
End If

If chkShowRejected Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'RJ'"
    Else
        xTransType = "'RJ'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'RJ'"
    Else
        xNotTransType = "'RJ'"
    End If
End If

If chkShowResubmitted Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'RS'"
    Else
        xTransType = "'RS'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'RS'"
    Else
        xNotTransType = "'RS'"
    End If
End If

If chkShowAppForward Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'AF'"
    Else
        xTransType = "'AF'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'AF'"
    Else
        xNotTransType = "'AF'"
    End If
End If

If chkShowAppFailed Then
    If Len(xTransType) > 0 Then
        xTransType = xTransType & ",'APF'"
    Else
        xTransType = "'APF'"
    End If
Else
    If Len(xNotTransType) > 0 Then
        xNotTransType = xNotTransType & ",'APF'"
    Else
        xNotTransType = "'APF'"
    End If
End If

If Len(xTransType) > 0 Then
    EECri = "({HR_REQAUDIT.RT_TRSTYPE} IN [" & xTransType & "])"

    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
Else
    If Len(xNotTransType) > 0 Then
        EECri = " NOT ({HR_REQAUDIT.RT_TRSTYPE} IN [" & xNotTransType & "])"
    
        If glbiOneWhere Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
    End If
End If

'If Len(xNotTransType) > 0 Then
'    EECri = "({HR_REQAUDIT.RT_TRSTYPE} NOT IN [" & xNotTransType & "])"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'End If

'If Not chkShowApproved Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'AP'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'If Not chkShowArchived Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'AR'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'
'If Not chkShowDeleted Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'DB'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'
'If Not chkShowDelApproved Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'DA'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'
'If Not chkShowModified Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'MB'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'
'If Not chkShowNewRequest Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'NR'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'
'If Not chkShowRejected Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'RJ'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If
'
'If Not chkShowResubmitted Then
'    EECri = "{HR_REQAUDIT.RT_TRSTYPE} <> 'RS'"
'
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'
'    glbiOneWhere = True
'End If

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_TransFTDates
Call Cri_ReqFTDates
Call Cri_VacTimeOffReq
Call Cri_TransactionTypes

' report name
If comGroup(0) <> "(none)" Then
    strRName$ = glbIHRREPORTS & "rzESSReqTrnAud.rpt"
Else
    strRName$ = glbIHRREPORTS & "rzESSReqTrnAud1.rpt"
End If
Me.vbxCrystal.ReportFileName = strRName$

'set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
  

Me.vbxCrystal.Connect = RptODBC_SQL

  
Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ESS Requests - Transaction Audit", "HR_REQAUDIT", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

Call setRptLabel(Me, 1)

If comGroup(0) = "(none)" Then
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        Case 2: grpField$ = "{HR_REQAUDIT.RT_LDATE}"
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ = "(none)" Then
        grpField$ = "{HR_REQAUDIT.RT_REQID}"
    End If
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    Exit Function
End If
    
Y% = X% + 1
dscGroup$ = comGroup(X%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(X%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    Case 2: grpField$ = "{HR_REQAUDIT.RT_LDATE}"
    Case 3: grpField$ = "(none)"
End Select
If grpField$ = "(none)" Then
    grpField$ = "{HR_REQAUDIT.RT_REQID}"
End If
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$


'GrpIdx% = comGroup(2).ListIndex
'Select Case GrpIdx%
'    Case 0: grpField$ = "{@EFullName}"
'    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
'End Select
'grpField$ = "{HR_REQAUDIT.RT_REQID}"
'grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(2) = grpCond$

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%
Dim dd As Long

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

If Not clpPT.ListChecker Then
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

'Request From/To Date
If Len(dlpReqFrom.Text) > 0 Then
    If Not IsDate(dlpReqFrom.Text) Then
        MsgBox "Invalid Request From date"
        dlpReqFrom.SetFocus
        Exit Function
    End If
End If
If Len(dlpReqTo.Text) > 0 Then
    If Not IsDate(dlpReqTo.Text) Then
        MsgBox "Invalid Request To date"
        dlpReqTo.SetFocus
        Exit Function
    End If
End If
If Len(dlpReqFrom.Text) > 0 And Len(dlpReqTo.Text) > 0 Then
    dd = DateDiff("d", CVDate(dlpReqFrom.Text), CVDate(dlpReqTo.Text))
    If dd < 0 Then
        MsgBox "Request From date must be earlier than Request To Date"
        dlpReqFrom.SetFocus
        Exit Function
    End If
End If

'Transaction From/To Date
If Len(dlpFrom.Text) > 0 Then
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid Transaction From date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Len(dlpTo.Text) > 0 Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid Transaction To date"
        dlpTo.SetFocus
        Exit Function
    End If
End If
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dd = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
    If dd < 0 Then
        MsgBox "Transaction From date must be earlier than Transaction To Date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If

CriCheck = True

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMRESSTRNAUDIT"

Screen.MousePointer = HOURGLASS
frmRESSTrnAudit.Caption = "ESS Requests - Transaction Audit Report"

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

lblGrp(3).Visible = False
comGroup(2).Visible = False
Call comGrpLoad

Call setRptCaption(Me)

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Resize()

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 10000 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 6000 Then
            scrControl.Max = 6000
        Else
            scrControl.Max = 3800
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 10000 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 50
        End If
        scrHScroll.Top = Me.Height - 700
        scrHScroll.Width = Me.Width - 120
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub scrControl_Change()
    scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
    scrFrame.Left = 40 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

