VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUATTEND 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Update Attendance"
   ClientHeight    =   11340
   ClientLeft      =   -210
   ClientTop       =   1350
   ClientWidth     =   10170
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11340
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pcSelectCri 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   120
      ScaleHeight     =   10455
      ScaleWidth      =   9615
      TabIndex        =   28
      Top             =   240
      Width           =   9615
      Begin VB.TextBox txtShift 
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
         Left            =   2360
         MaxLength       =   4
         TabIndex        =   19
         Tag             =   "00-Shift code"
         Top             =   7470
         Width           =   435
      End
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
         Height          =   930
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Tag             =   "00-Comments - free form Memo field"
         Top             =   8340
         Width           =   8790
      End
      Begin VB.TextBox txtPoint 
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
         Left            =   2360
         MaxLength       =   20
         TabIndex        =   20
         Tag             =   "00-Point"
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Frame frAttachment 
         Caption         =   "For Mass Add only"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   6000
         TabIndex        =   30
         Top             =   9480
         Width           =   3015
         Begin VB.CommandButton cmdImport 
            Caption         =   "Import"
            Height          =   270
            Left            =   1785
            TabIndex        =   27
            Top             =   360
            Width           =   855
         End
         Begin VB.Image imgSec 
            Height          =   240
            Left            =   1365
            Picture         =   "fuattnd.frx":0000
            Top             =   375
            Width           =   240
         End
         Begin VB.Image imgNoSec 
            Height          =   240
            Left            =   1365
            Picture         =   "fuattnd.frx":014A
            Top             =   375
            Width           =   240
         End
         Begin VB.Label lblImport 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Attendance"
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
            Left            =   240
            TabIndex        =   31
            Top             =   375
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   120
         TabIndex        =   29
         Top             =   280
         Width           =   5295
         Begin VB.OptionButton optEmployee 
            Alignment       =   1  'Right Justify
            Caption         =   "Active Employees"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   150
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optEmployee 
            Alignment       =   1  'Right Justify
            Caption         =   "Terminated Employees"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   1
            Top             =   150
            Width           =   2535
         End
      End
      Begin INFOHR_Controls.EmployeeLookup elpSup 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Tag             =   "11-Employee Number of individual's supervisor"
         Top             =   6810
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   1875
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin MSMask.MaskEdBox medHours 
         Height          =   285
         Left            =   2355
         TabIndex        =   18
         Tag             =   "11-Hours for this reason (> 0)"
         Top             =   7140
         Width           =   1065
         _ExtentX        =   1879
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
      Begin Threed.SSCheck chkIncentive 
         Height          =   195
         Left            =   4875
         TabIndex        =   21
         Tag             =   "Incentive -  Attendance Management"
         Top             =   7515
         Width           =   1320
         _Version        =   65536
         _ExtentX        =   2328
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "Incentive"
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
      Begin Threed.SSCheck chkSeniority 
         Height          =   225
         Left            =   4875
         TabIndex        =   23
         Tag             =   "Hours to be added to employee's seniority."
         Top             =   7830
         Width           =   960
         _Version        =   65536
         _ExtentX        =   1693
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Seniority"
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
      Begin Threed.SSCheck chkIncident 
         Height          =   255
         Left            =   6180
         TabIndex        =   22
         Tag             =   "Is this a new incidence of illness?"
         Top             =   7485
         Width           =   1065
         _Version        =   65536
         _ExtentX        =   1879
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Incident"
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
      Begin Threed.SSCheck chkEMELEA 
         Height          =   225
         Left            =   6180
         TabIndex        =   24
         Tag             =   "Emergency Leave"
         Top             =   7830
         Width           =   1560
         _Version        =   65536
         _ExtentX        =   2752
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Emergency Leave"
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
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Tag             =   "00-Specific Department Desired"
         Top             =   1185
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
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Tag             =   "00-Specific Division Desired"
         Top             =   840
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
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "00-Enter Union Code"
         Top             =   1530
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDOR"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Tag             =   "EDPT-Category"
         Top             =   2220
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
         Index           =   5
         Left            =   1920
         TabIndex        =   7
         Tag             =   "00-Enter Region Code"
         Top             =   2565
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDRG"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   14
         Tag             =   "01-Attendance/absentee Reason"
         Top             =   5820
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ADRE"
      End
      Begin INFOHR_Controls.DateLookup dlpToDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   16
         Tag             =   "40-Date upto and including this date forward"
         Top             =   6480
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpAttDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Tag             =   "40-Date from and including this date forward"
         Top             =   6150
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Tag             =   "10-Enter Employee Number"
         Top             =   2910
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   9
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   3255
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDLC"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Tag             =   "00-Specific Employment Status Desired"
         Top             =   3615
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDSE"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   5040
         Top             =   9720
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   1920
         TabIndex        =   11
         Top             =   3960
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDAB"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3630
         TabIndex        =   13
         Tag             =   "40-Date upto and including this date forward"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Tag             =   "40-Date from and including this date forward"
         Top             =   4320
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
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
         TabIndex        =   53
         Top             =   885
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
         TabIndex        =   52
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Union Code"
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
         Top             =   1575
         Width           =   840
      End
      Begin VB.Label lblTitle 
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
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label textMulti 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "The Union Code and Category will be validated from the Employee Basic Data"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   49
         Top             =   5040
         Visible         =   0   'False
         Width           =   7455
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   48
         Top             =   5865
         Width           =   660
      End
      Begin VB.Label lblAttendance 
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance"
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
         Top             =   5520
         Width           =   1215
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
         TabIndex        =   46
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
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
         Left            =   180
         TabIndex        =   45
         Top             =   6195
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hours"
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
         Left            =   180
         TabIndex        =   44
         Top             =   7170
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Index           =   8
         Left            =   180
         TabIndex        =   43
         Top             =   7485
         Width           =   315
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Supervisor"
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
         TabIndex        =   42
         Top             =   6840
         Width           =   750
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         Index           =   10
         Left            =   180
         TabIndex        =   41
         Top             =   8115
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
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
         Left            =   150
         TabIndex        =   40
         Top             =   6510
         Width           =   585
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
         TabIndex        =   39
         Top             =   2610
         Width           =   510
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
         TabIndex        =   38
         Top             =   2955
         Width           =   1290
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
         TabIndex        =   37
         Top             =   2265
         Width           =   630
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
         TabIndex        =   36
         Top             =   3300
         Width           =   615
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   35
         Top             =   3660
         Width           =   540
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Point"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   180
         TabIndex        =   34
         Top             =   7815
         Width           =   1215
      End
      Begin VB.Label lblAdmin 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   33
         Top             =   4005
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Term'n. Date Range"
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
         TabIndex        =   32
         Top             =   4365
         Visible         =   0   'False
         Width           =   1440
      End
   End
   Begin VB.VScrollBar scrControl 
      Height          =   10215
      LargeChange     =   315
      Left            =   9820
      Max             =   100
      SmallChange     =   315
      TabIndex        =   26
      Top             =   240
      Width           =   340
   End
End
Attribute VB_Name = "frmUATTEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbAdd%
Dim fglbDelete%
Dim fglbModify%
Dim fglbSDate As Variant
Dim fglbESQLQ, fglbWSQLQ
Dim xASL As String
Dim xDiscipFlag As Boolean, xOccuAmount
Dim yDiscipFlag As Boolean
Dim fglbRetry, xmedHours, xAnother
Dim SavEML, SavVac, SavSick, AddChg, cntSick, savIncid, SavOutE, SavOutV, SavOutS, SaveHours
Dim Fdate, Tdate, fdateS, tdateS
Dim strEMPLIST, strTERMSEQ 'George Mar 14,2006
Dim RSEMPLIST As New ADODB.Recordset 'George Mar 14,2006
Dim xKey
Dim flgSkipESSAP As Boolean

Private Sub ChkEMELEA_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkIncentive_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkIncident_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function chkMUAttend(xUpdType)

Dim SQLQ As String, Msg$, dd&, Response%, X%
Dim DgDef As Variant, Title$, DCurPDate As Variant

chkMUAttend = False

On Error GoTo chkMUAttend_Err


If optEmployee(1) And (xUpdType = "A" Or xUpdType = "M") Then  'Release 8.0 - Only Delete Attendance for Terminated Employees
    MsgBox "You cannot 'Mass Add' or 'Mass Update' for 'Terminated Employees'. You can only 'Mass Delete' for 'Terminated Employees'", vbExclamation
    optEmployee(1).SetFocus
    Exit Function
End If

'Ticket #26576 - WDGPHU - Cannot add FX* codes from Attendance
If glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC And (xUpdType = "A" Or xUpdType = "M" Or xUpdType = "D") And UCase(Left(clpCode(4), 2)) = "FX" Then
    MsgBox "You cannot 'Mass Add' or 'Mass Update' or 'Mass Delete' Flex Time Attendance from here. Please use ESS Module.", vbExclamation, "info:HR - Flex Time entry restricted"
    Exit Function
End If

'Ticket #30305 - Disable Compensatory Time Entries
If gsDISABLE_COMPTIME Then
    If Left(clpCode(4).Text, 2) = "OT" Or Left(clpCode(4).Text, 2) = "CT" Then
        MsgBox "You cannot 'Mass Add' or 'Mass Update' or 'Mass Delete' Compensatory Time Attendance from here. Please use ESS Module.", vbExclamation, "info:HR - Compensatory Time entry restricted"
        Exit Function
    End If
End If

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be known")
'    clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox "If Department Entered - it must be known"
'     clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
    'If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    '    MsgBox "If code entered it must be known"
    '    clpCode(X%).SetFocus
    '    Exit Function
    'End If
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
'    MsgBox lStr("Category code must be valid")
'    clpPT.SetFocus
    Exit Function
End If

If Not optEmployee(0) And Not optEmployee(1) Then
    MsgBox "The type of employees 'Active Employees' or 'Terminated Employees' must be selected"
    optEmployee(0).SetFocus
    Exit Function
End If

If optEmployee(1) Then
    For X% = 0 To 1
        If Len(dlpDateRange(X%).Text) > 0 Then
            If Not IsDate(dlpDateRange(X%).Text) Then
                MsgBox "Not a valid Termination Date"
                dlpDateRange(X%).Text = ""
                dlpDateRange(X%).SetFocus
                Exit Function
            End If
        End If
    Next X%
    
    If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
        If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then
            MsgBox "Termination To Date can't be prior to Termination From Date!"
            Me.dlpDateRange(0).SetFocus
            Exit Function
        End If
    End If
End If

If Len(clpCode(4).Text) < 1 Then
    MsgBox "Attendance Reason is a required field"
    clpCode(4).SetFocus
    Exit Function
End If

If Len(dlpAttDate.Text) < 1 Then
    If Not fglbDelete Then
        Msg$ = "Enter Attendance Date!"
        dlpAttDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If
Else
    If Not IsDate(dlpAttDate.Text) Then
        Msg$ = "Not a valid Attendance From Date"
        dlpAttDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If
If Len(dlpToDate.Text) > 0 Then
    If Not IsDate(dlpToDate.Text) Then
        Msg$ = "Not a valid Attendance To Date"
        dlpToDate.SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If

If IsDate(dlpAttDate.Text) And IsDate(dlpToDate.Text) Then
    If DaysBetween(dlpAttDate, dlpToDate) < 0 Then
        MsgBox "Attendance To Date can't be prior to Attendance From Date!"
        Me.dlpAttDate.SetFocus
        Exit Function
    End If
End If


If fglbDelete Then GoTo chkMUOK        ' rest not checked if delete.

If Len(medHours) > 0 Then
    If Not IsNumeric(medHours) Then
        MsgBox "Hours is invalid"
        medHours.SetFocus
        Exit Function
    End If
Else
    medHours = 0
End If

'Ticket #15323
'If glbLinamar Then
'    If Len(txtShift) < 1 Then
'        MsgBox "Shift is a required field"
'        txtShift.SetFocus
'        Exit Function
'    End If
'End If

If Len(elpSup.Text) > 0 And elpSup.Caption = "Unassigned" Then
    MsgBox "Supervisor is invalid"
    elpSup.SetFocus
    Exit Function
End If

chkMUOK:
If Not elpEEID.ListChecker Then
    Exit Function
End If



chkMUAttend = True

Exit Function

chkMUAttend_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMUAttend", "HR Attendance", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub chkSeniority_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, X%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Integer

If Not gSec_Upd_Attendance Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

strEMPLIST = ""

fglbDelete% = True
fglbAdd% = False
fglbModify% = False

If Not chkMUAttend("D") Then Exit Sub

Title$ = "Mass Attendance Records Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If optEmployee(0) Then Msg$ = Msg$ & " Active Employees" Else Msg$ = Msg$ & " Terminated Employees"
    If recCount = 1 Then Msg$ = Msg$ & " Attendance Record " Else Msg$ = Msg$ & " Attendance Records "
    Msg$ = Msg$ & "to Delete. This delete is not reversible." & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Attendance records found to delete."
    GoTo End_Note
End If

If optEmployee(0) Then  'Release 8.0
    Msg$ = "Do you want to print a list of Active Employees updated?"
Else
    Msg$ = "Do you want to print a list of Terminated Employees employees updated?"
End If
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Not modDelRecs() Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

If Len(strEMPLIST) > 0 Then
    MsgBox "Records Deleted Successfully."
Else
    MsgBox "0 Records Deleted."
End If

If Response% = IDYES Then    ' Yes response
    'Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    
    'Call getWSQLQ("U")
    
    'report name
    If optEmployee(0) Then  'Release 8.0
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Delete Attendance - Active Employees Details'"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZTEmpList.rpt"
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Delete Attendance - Terminated Employees Details'"
    End If
    
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        If Len(strEMPLIST) > 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        Else
            Me.vbxCrystal.SelectionFormula = "1=2"
        End If
    End If
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '            Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    
    ' window title if appropriate
    If optEmployee(0) Then  'Release 8.0
        Me.vbxCrystal.WindowTitle = "Employees-updated Report"
    Else
        Me.vbxCrystal.WindowTitle = "Terminated Employees-updated Report"
    End If
    
    Me.vbxCrystal.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
End If

End_Note:

Screen.MousePointer = DEFAULT

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
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%
Dim recCount As Integer

On Error GoTo Mod_Err

If Not gSec_Upd_Attendance Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

strEMPLIST = ""
fglbDelete% = False
fglbAdd% = False
fglbModify% = True

If Not chkMUAttend("M") Then Exit Sub

Title$ = "Mass Update Attendance"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to update all Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Update
If recCount > 0 Then
    Msg$ = Str(recCount)
    If optEmployee(0) Then Msg$ = Msg$ & " Active Employees" Else Msg$ = Msg$ & " Terminated Employees"
    If recCount = 1 Then Msg$ = Msg$ & " Attendance Record " Else Msg$ = Msg$ & " Attendance Records "
    Msg$ = Msg$ & "to Update. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Attendance records found to update."
    GoTo End_Note
End If

If optEmployee(0) Then  'Release 8.0
    Msg$ = "Do you want to print a list of Active Employees updated?"
Else
    Msg$ = "Do you want to print a list of Terminated Employees employees updated?"
End If
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Not modUpdRecs() Then Exit Sub

If Len(strEMPLIST) > 0 Then
    MsgBox "Records Updated Successfully."
Else
    MsgBox "0 Records Updated."
End If

If Response% = IDYES Then    ' Yes response
    'Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    
    'Call getWSQLQ("U")
    
    'report name
    If optEmployee(0) Then  'Release 8.0
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Attendance - Active Employees Details'"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZTEmpList.rpt"
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Attendance - Terminated Employees Details'"
    End If
    
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        If Len(strEMPLIST) > 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        Else
            Me.vbxCrystal.SelectionFormula = "1=2"
        End If
    End If
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '            Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    
    ' window title if appropriate
    If optEmployee(0) Then  'Release 8.0
        Me.vbxCrystal.WindowTitle = "Active Employees-updated Report"
    Else
        Me.vbxCrystal.WindowTitle = "Terminated Employees-updated Report"
    End If
    
    Me.vbxCrystal.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
End If

End_Note:
Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String, Msg$, X%
Dim Title$, DgDef As Variant, Response%
Dim recCount As Integer

On Error GoTo AddN_Err

strEMPLIST = ""

If Not gSec_Upd_Attendance Then
  'tkt310423 Jerry said remove serial#control for Add_Attendance security
  '  If glbCompSerial = "S/N - 2173W" Then   'Ticket #7500 - Town of Ajax
        If Not gSec_Add_Attendance Then
            MsgBox "You Do Not Have Authority For This Transaction"
            Exit Sub
        End If
'    Else
'        MsgBox "You Do Not Have Authority For This Transaction"
'        Exit Sub
   ' End If
End If

fglbAdd% = True
fglbDelete% = False
fglbModify% = False

If Not chkMUAttend("A") Then Exit Sub

Title$ = "Mass Records Attendance"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to add Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Add
If recCount > 0 Then
    Msg$ = Str(recCount)
    If optEmployee(0) Then Msg$ = Msg$ & " Active Employees" Else Msg$ = Msg$ & " Terminated Employees"
    If recCount = 1 Then Msg$ = Msg$ & " Attendance Record " Else Msg$ = Msg$ & " Attendance Records "
    Msg$ = Msg$ & "will be Added. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Employee record found to add the Attendance record."
    GoTo End_Note
End If

If optEmployee(0) Then  'Release 8.0
    Msg$ = "Do you want to print a list of Active Employees updated?"
Else
    Msg$ = "Do you want to print a list of Terminated Employees employees updated?"
End If
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Not modInsRecs() Then Exit Sub

If Len(strEMPLIST) > 0 Then
    MsgBox "Records Added Successfully"
Else
    MsgBox "0 Records Added"
End If

If Response% = IDYES Then    ' Yes response
    'Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    
    'Call getWSQLQ("U")
    
    ' report name
    If optEmployee(0) Then  'Release 8.0
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList.rpt"
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Attendance - Active Employees Details'"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZTEmpList.rpt"
        Me.vbxCrystal.Formulas(0) = "rTitle='Mass Update Attendance - Terminated Employees Details'"
    End If
    
    'set location for database tables
    If Len(glbstrSelCri) >= 0 Then
        If Len(strEMPLIST) > 0 Then
            Me.vbxCrystal.SelectionFormula = getWSQLQRPT
        Else
            Me.vbxCrystal.SelectionFormula = "1=2"
        End If
    End If
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '            Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    
    ' window title if appropriate
    If optEmployee(0) Then  'Release 8.0
        Me.vbxCrystal.WindowTitle = "Active Employees-updated Report"
    Else
        Me.vbxCrystal.WindowTitle = "Terminated Employees-updated Report"
    End If
    
    Me.vbxCrystal.Destination = 0
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
End If

End_Note:

Screen.MousePointer = DEFAULT

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
Resume Next
End Sub

Private Sub ATTCode_Desc(Indx As Integer)
Dim SQLQ As String
Dim rsCode As New ADODB.Recordset
On Error GoTo ATTCode_Err

If Not Indx = 4 Then Exit Sub
If Len(clpCode(Indx).Text) > 0 Then
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' "
    SQLQ = SQLQ & " AND TB_KEY='" & clpCode(Indx).Text & "'"
    
    rsCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    chkEMELEA.Enabled = False
    If Not rsCode.EOF Then
        If Indx = 4 Then
            If rsCode("TB_USR3") <> 0 Then
                chkEMELEA.Enabled = True
            End If
            chkIncentive.Value = rsCode("TB_INDICATOR")
            chkSeniority.Value = rsCode("TB_SEN")
            chkEMELEA.Value = rsCode("TB_USR3")
            If IsNumeric(rsCode("TB_USR2")) Then
                txtPoint = rsCode("TB_USR2")
            Else
                txtPoint = ""
            End If
        End If
    End If
End If
Exit Sub

ATTCode_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ATT Code Snap", "Key", "SELECT")
Resume Next

End Sub

Private Sub clpCode_Change(Index As Integer)
If Index = 4 Then Call ATTCode_Desc(Index)
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = True
    glbDocName = "AttendanceMU"
    
    If glbDocNewRecord Then
        glbDocKey = 0
        
        If Len(dlpAttDate.Text) = 0 Or Len(clpCode(4).Text) = 0 Then
            MsgBox "'" & lStr("From Date") & "' and '" & lStr("Reason") & "' must be entered before attaching a document.", vbExclamation
            Exit Sub
        Else
            glbAttReason = clpCode(4).Text
            glbAttDOA = dlpAttDate.Text
        End If
    End If
    frmInAttachment.Show 1
    DoEvents
    
    glbDocName = "Attendance"
    glbLEE_ID = 0
    
    Call DispimgIcon(Me, "frmUATTEND")
    
    If glbDocImpFile <> "" Then
        imgSec.Visible = True
        imgNoSec.Visible = False
    Else
        imgSec.Visible = False
        imgNoSec.Visible = True
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUATTEND"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

glbOnTop = "FRMUATTEND"

Call setRptCaption(Me)

If glbLinamar Then
    lblRegion.Visible = True
    clpCode(5).Visible = True
    clpCode(5).MaxLength = 8
End If
If glbCompSerial = "S/N - 2227W" Then
    clpCode(5).MaxLength = 6
End If
If glbCompSerial = "S/N - 2347W" Then 'Surrey Place
    chkIncentive.Caption = "LTD"
    chkIncentive.Tag = "LTD"
End If
If glbCompSerial = "S/N - 2388W" Then 'DNSSAB Ticket #14260
    chkIncentive.Caption = "No Sick Ent"
    chkIncentive.Tag = "No Sick Ent"
End If
'Ticket #15323
'If glbLinamar Then
'    lblTitle(8).FontBold = True
'End If
If glbBurlTech Then
    chkIncident.Visible = False
    chkIncentive.Caption = "Unexcused"
    chkSeniority.Caption = "Excused"
End If

'WDGPHU - Ticket #26576
'Leeds and Grenville - Ticket #19441
If glbCompSerial = "S/N - 2233W" Or (glbCompSerial = "S/N - 2411W" And gsFLEX_LOGIC) Then
    chkIncentive.Caption = "Frozen"
    chkIncentive.Tag = "Frozen"
    chkIncentive.Enabled = False
End If
If glbMulti Then textMulti.Visible = True
textMulti.Caption = "The " & lStr("Union") & " and " & lStr("Category") & " will be validated from the Employee Basic Data"

Call INI_Controls(Me)

'Initializing the values
glbDocImpFile = ""
glbDocType = ""
glbDocDesc = ""

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Resize()
Dim c As Long

On Error GoTo Eh

If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    If Me.Height >= 10500 Then
        scrControl.Value = 0
        
        pcSelectCri.Top = 120
        
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        scrControl.Left = Me.ScaleWidth - scrControl.Width
        scrControl.Height = Me.Height - 1000
        
        scrControl.Max = 5000
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
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Form_Resize", "Attendance Mass Update", "Resize")
    Resume exH

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmUATTEND = Nothing 'carmen apr 2000
End Sub

Private Sub imgSec_Click()
    If Len(glbDocImpFile) > 0 Then
        Shell "cmd /c " & GetShortName(glbDocImpFile)
    End If
End Sub

Private Sub medHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(ActiveControl)
MDIMain.panHelp(2).Caption = " "
End Sub

Private Function modDelRecs()
Dim BD As Integer
Dim SQLQ As String, SQL1 As String, countr As Integer, WSQLQ
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$
Dim chkSQLQ As String     'Laura
Dim dynHRAT As New ADODB.Recordset
Dim TSQLQ
Dim rsAttDel As New ADODB.Recordset
Dim xDayNum As Integer, xInt As Integer, xtmpdate
Dim DgDef, Title$, Msg$, Response%
Dim xExclude As Boolean
Dim xlocReason, xlocDate

modDelRecs = False

On Error GoTo modDelRecs_Err

Screen.MousePointer = HOURGLASS

xExclude = False

Call getWSQLQ

WSQLQ = WSQLQ & " AD_REASON = '" & clpCode(4).Text & "' "

If Len(dlpToDate.Text) > 0 Then
    If Len(dlpAttDate.Text) > 0 Then
        WSQLQ = WSQLQ & " AND AD_DOA >= " & Date_SQL(dlpAttDate.Text)
    End If
    WSQLQ = WSQLQ & " AND AD_DOA <= " & Date_SQL(dlpToDate.Text)
Else
    If Len(dlpAttDate.Text) > 0 Then
        WSQLQ = WSQLQ & " AND AD_DOA = " & Date_SQL(dlpAttDate.Text)
    End If
End If

If Len(medHours.Text) > 0 Then
    WSQLQ = WSQLQ & " AND AD_HRS = " & medHours.Text
End If

If glbCompSerial = "S/N - 2192W" Then   'County of Essex
    'Exclude Machine # populated records
    If Len(medHours.Text) > 0 Then 'Ticket #12338
        If medHours.Text = 0 Then
            WSQLQ = WSQLQ & " AND (AD_MACHINE_NUM IS NULL OR AD_MACHINE_NUM = '')"
        End If
    End If
End If

If optEmployee(0) Then  'Release 8.0
    SQLQ = "SELECT AD_ATT_ID, AD_EMPNBR, AD_REASON, AD_DOA, AD_DOCKEY, AD_SOURCE, AD_REQID FROM HR_ATTENDANCE WHERE " & WSQLQ
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
    
    'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted - Warn the users
    'Check if there are any ESS Approved records in this selection criteria
    If ContainsESSApprovedAttendance(SQLQ) Then
        Msg$ = "There are some ESS Approved Attendance records in this selection." & vbCrLf & vbCrLf & "Do you want exclude them from Delete?"
        Response% = MsgBox(Msg$, vbExclamation + vbYesNoCancel, "ESS Approved Attendance Records found")     ' Get user response.
        If Response = IDCANCEL Then
            Screen.MousePointer = DEFAULT
            Exit Function
        ElseIf Response = IDYES Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
            xExclude = True
        End If
    End If
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT AD_ATT_ID, AD_EMPNBR, AD_REASON, AD_DOA, AD_DOCKEY, AD_SOURCE, AD_REQID  FROM TERM_ATTENDANCE WHERE " & WSQLQ
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM TERM_HREMP WHERE " & fglbESQLQ & ")"
                
    'Termination Date Range
    If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
        SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
        'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
        SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0))
        SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
    Else
        If IsDate(dlpDateRange(0)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0)) & ")"
        End If
        If IsDate(dlpDateRange(1)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
        End If
    End If
        
    'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted - Warn the users
    'Check if there are any ESS Approved records in this selection criteria
    If ContainsESSApprovedAttendance(SQLQ) Then
        Msg$ = "There are some ESS Approved Attendance records in this selection." & vbCrLf & vbCrLf & "Do you want exclude them from Delete?"
        Response% = MsgBox(Msg$, vbExclamation + vbYesNoCancel, "ESS Approved Attendance Records found")     ' Get user response.
        If Response = IDCANCEL Then
            Screen.MousePointer = DEFAULT
            Exit Function
        ElseIf Response = IDYES Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
            xExclude = True
        End If
    End If
    dynHRAT.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
End If
If dynHRAT.BOF And dynHRAT.EOF Then
    modDelRecs = False
    MsgBox "Records Selection Not Found!"
    Exit Function
Else
    'If glbVadim Then
     '   Do Until dynHRAT.EOF
    '        dynHRAT.Delete
    '        dynHRAT.MoveNext
    '    Loop
    'Else
    If optEmployee(0) Then  'Release 8.0
        SQLQ = "SELECT DISTINCT AD_EMPNBR FROM HR_ATTENDANCE WHERE " & WSQLQ
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
        
        'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted
        If xExclude Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
        End If
        
        RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    Else
        SQLQ = "SELECT DISTINCT AD_EMPNBR, TERM_SEQ FROM TERM_ATTENDANCE WHERE " & WSQLQ
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM TERM_HREMP WHERE " & fglbESQLQ & ")"
        
        'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted
        If xExclude Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
        End If
        
        'Termination Date Range
        If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0))
            SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
        Else
            If IsDate(dlpDateRange(0)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0)) & ")"
            End If
            If IsDate(dlpDateRange(1)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
            End If
        End If
        
        RSEMPLIST.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly, adLockReadOnly
        strTERMSEQ = ""
    End If
    Do While Not RSEMPLIST.EOF
        If Len(strEMPLIST) > 0 Then
            strEMPLIST = strEMPLIST & "," & RSEMPLIST("AD_EMPNBR")
            
            If optEmployee(1) Then  'Release 8.0
                strTERMSEQ = strTERMSEQ & "," & RSEMPLIST("TERM_SEQ")
            End If
        Else
            strEMPLIST = strEMPLIST & RSEMPLIST("AD_EMPNBR")
            
            If optEmployee(1) Then  'Release 8.0
                strTERMSEQ = strTERMSEQ & RSEMPLIST("TERM_SEQ")
            End If
        End If
        RSEMPLIST.MoveNext
    Loop
    RSEMPLIST.Close

    '7.9 Enhancement - Delete the Attachment documents if any documents attached
    If gsAttachment_DB Then
        glbDocName = "Attendance"
        
        'Delete document for each record
        dynHRAT.MoveFirst
        Do While Not dynHRAT.EOF
            If Not IsNull(dynHRAT("AD_DOCKEY")) Then
                If dynHRAT("AD_DOCKEY") <> "" Then
                    glbDocKey = dynHRAT("AD_DOCKEY")
                    If optEmployee(0) Then  'Release 8.0
                        gdbAdoIhr001_DOC.Execute "DELETE FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR = " & dynHRAT("AD_EMPNBR") & " AND AD_REASON='" & clpCode(4).Text & "' AND AD_DOA=" & Date_SQL(dynHRAT("AD_DOA")) & " and AD_DOCKEY=" & glbDocKey & " "
                    Else
                        gdbAdoIhr001_DOC.Execute "DELETE FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR = " & dynHRAT("AD_EMPNBR") & " AND AD_REASON='" & clpCode(4).Text & "' AND AD_DOA=" & Date_SQL(dynHRAT("AD_DOA")) & " and AD_DOCKEY=" & glbDocKey & " "
                    End If
                End If
            End If
            dynHRAT.MoveNext
        Loop
    End If
    
    If optEmployee(0) Then  'Release 8.0
        SQLQ = "DELETE FROM HR_ATTENDANCE WHERE " & WSQLQ
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
        
        'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted
        If xExclude Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
        End If
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute SQLQ
        gdbAdoIhr001.CommitTrans
    Else
        SQLQ = "DELETE FROM Term_ATTENDANCE WHERE " & WSQLQ
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & ")"
        
        'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted
        If xExclude Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
        End If
        
        'Termination Date Range
        If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0))
            SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
        Else
            If IsDate(dlpDateRange(0)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0)) & ")"
            End If
            If IsDate(dlpDateRange(1)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
            End If
        End If
        
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute SQLQ
        gdbAdoIhr001X.CommitTrans
    End If
    
    'End If
    modDelRecs = True
End If
dynHRAT.Close
Set dynHRAT = Nothing

'If glbBurlTech Then 'BTI Points Recalculate
'    Call BTIPoint(fglbESQLQ)
'End If

'Ticket #12718
'If glbAdv Then
If glbAdv Or glbWFC Then 'Ticket #28919 Franks 03/24/2017
    If IsDate(dlpAttDate.Text) Then
        SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ
        rsAttDel.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Do While Not rsAttDel.EOF
            xDayNum = 0
            xtmpdate = dlpAttDate.Text
            If IsDate(dlpToDate.Text) Then
                xDayNum = DateDiff("D", dlpAttDate.Text, dlpToDate.Text)
            End If
            If xDayNum < 0 Then xDayNum = 0
            For xInt = 0 To xDayNum
                xKey = rsAttDel("ED_EMPNBR")
                xKey = xKey & "|" & Format(xtmpdate, "dd-mmm-yyyy")
                xKey = xKey & "|" & clpCode(4).Text
                Call Attendance_Master_Integration(xKey, , True)
                If glbWFC Then  'Ticket #28919 Franks 03/24/2017
                    Call WFC_Attend_To_AT(rsAttDel("ED_EMPNBR"), "D", xtmpdate, clpCode(4).Text, 0, , , "U")
                End If
                
                xtmpdate = DateAdd("D", 1, xtmpdate)

    
            Next
            rsAttDel.MoveNext
        Loop
    End If
End If

If glbCompSerial <> "S/N - 2192W" Then   'County of Essex - Because they want to manually calculate employee's vac anniversary date
                                         ' and they just trying to delete Regular 0 hours records
    Call EntReCalc(fglbESQLQ)
End If
Call EntReCalcHr

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    Call ReCalcOvt("")
'End If

'Town of Ajax
If glbCompSerial = "S/N - 2173W" Then
    Call Recalculate_OTBANK
End If

modDelRecs = True

Exit Function

modDelRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "DeleteAttend", "Delete")
modDelRecs = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modInsRecs()
Dim SQLQ As String
Dim dyn_HREDesem As New ADODB.Recordset, rsDup As New ADODB.Recordset
Dim fblDoul, xDays
Dim X, xDup
Dim xDATE
Dim WSQLQ, ESQLQ, Result
Dim TSQLQ
Dim Msg$
Dim AskWeekend, SkipWeekend, AskHoliday, SkipHoliday
Dim xWeekDay
Dim Title$, DgDef As Variant, Response%
Dim rsAttSal As New ADODB.Recordset, rsCurSal As New ADODB.Recordset
Dim rsempt As New ADODB.Recordset, rsAttD As New ADODB.Recordset
Dim hlist As String

modInsRecs = False
On Error GoTo modInsRecs_Err


Screen.MousePointer = HOURGLASS

Call getWSQLQ

If optEmployee(0) Then  'Release 8.0
    ESQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ
    dyn_HREDesem.Open ESQLQ, gdbAdoIhr001, adOpenKeyset
Else
    ESQLQ = "SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ
    dyn_HREDesem.Open ESQLQ, gdbAdoIhr001X, adOpenKeyset
End If

If dyn_HREDesem.EOF And dyn_HREDesem.BOF Then
    modInsRecs = False
    MsgBox "Records for this selection do not exist!"
    Screen.MousePointer = DEFAULT
    Exit Function
End If
dyn_HREDesem.Close

'Surrey Place to check if the Employee Status is "TERM"
If glbCompSerial = "S/N - 2347W" And optEmployee(0) Then    'Release 8.0
    Dim xFLAG As Boolean
    
    SQLQ = ESQLQ '= "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ
    SQLQ = SQLQ & " AND ED_EMP = 'TERM'"
    xFLAG = False
    dyn_HREDesem.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    If Not dyn_HREDesem.EOF Then
        xFLAG = True
    End If
    dyn_HREDesem.Close
    If xFLAG Then
        Title$ = "Terminated Employee"
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = "The selection criteria includes terminated employees" & Chr(10)
        Msg$ = Msg$ & "Do you wish to add attendance records? " & Chr(10)
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
    End If
End If

fblDoul = CDbl(medHours)
If UCase(clpCode(4).Text) = "OT15" Then fblDoul = fblDoul * 1.5
If UCase(clpCode(4).Text) = "OT20" Then fblDoul = fblDoul * 2

'City of Timmins - Ticket #16168
If glbCompSerial = "S/N - 2375W" Then
    If UCase(clpCode(4).Text) = "OT05" Then fblDoul = fblDoul * 0.5
    If UCase(clpCode(4).Text) = "OT25" Then fblDoul = fblDoul * 2.5
End If

If Len(dlpToDate.Text) = 0 Then
    xDays = 0
Else
    xDays = DateDiff("d", dlpAttDate.Text, dlpToDate.Text)
End If
xDATE = dlpAttDate.Text

AskWeekend = True
AskHoliday = True

For X = 0 To xDays
   xWeekDay = Weekday(xDATE)
   If xWeekDay = 7 Or xWeekDay = 1 Then
        If AskWeekend Then
            Msg$ = "Do you want exclude Saturday/Sunday?"
            AskWeekend = False
            SkipWeekend = False
            If MsgBox(Msg$, 36) = 6 Then
                SkipWeekend = True
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
        Else
            If SkipWeekend Then
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
        End If
    End If
    If Len(dlpToDate.Text) > 0 Then
        If CVDate(xDATE) > CVDate(dlpToDate.Text) Then Exit For
    Else
        If CVDate(xDATE) > CVDate(dlpAttDate.Text) Then Exit For
    End If
    
    If optEmployee(0) Then  'Release 8.0
        TSQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE "
        TSQLQ = TSQLQ & " WHERE AD_REASON = '" & clpCode(4).Text & "' "
        TSQLQ = TSQLQ & " AND AD_DOA = " & Date_SQL(xDATE)
        TSQLQ = TSQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
        rsDup.Open TSQLQ, gdbAdoIhr001, adOpenKeyset
    Else
        TSQLQ = "SELECT AD_EMPNBR FROM Term_ATTENDANCE "
        TSQLQ = TSQLQ & " WHERE AD_REASON = '" & clpCode(4).Text & "' "
        TSQLQ = TSQLQ & " AND AD_DOA = " & Date_SQL(xDATE)
        TSQLQ = TSQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & ")"
        rsDup.Open TSQLQ, gdbAdoIhr001X, adOpenKeyset
    End If
    If Not rsDup.EOF Then
        Msg$ = "Reason: " & clpCode(4) & Chr(10) & " Date: " & xDATE & Chr(10) & Chr(10)
        Msg$ = Msg$ & rsDup.RecordCount & " duplicates found in Attendance Master. " & Chr(10) & Chr(10)
        Msg$ = Msg$ & "Click Yes to post all Attendance records including duplicates." & Chr(10)
        Msg$ = Msg$ & "Click No to post all non-duplicate Attendance records." & Chr(10)
        Result = MsgBox(Msg$, vbYesNo, "Duplicates Found")
        If Result = vbYes Then
            xDup = False
        Else
            xDup = True
        End If
    End If
    rsDup.Close
    
    If Not (glbWFC And glbPlantCode = "WHBY") Then
        'For VADIM
        'Open a empty record of Attendance table
        If optEmployee(0) Then  'Release 8.0
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_REASON ='**'"
            rsAttD.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO FROM HREMP WHERE" & fglbESQLQ
            If xDup Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (" & TSQLQ & ")"
            rsempt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Else
            SQLQ = "SELECT * FROM Term_ATTENDANCE WHERE AD_REASON ='**'"
            rsAttD.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            
            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO FROM Term_HREMP WHERE" & fglbESQLQ
            If xDup Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (" & TSQLQ & ")"
            rsempt.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        End If
                
        Do While Not rsempt.EOF
            'Check if Holidays should be excluded
            If AskHoliday Then
                hlist = IsSTATHoliday(CVDate(xDATE), CVDate(xDATE), rsempt("ED_EMPNBR"))
                If InStr(hlist, Date_SQL(xDATE)) > 0 Then
                    Msg$ = "Do you want exclude STAT holidays?"
                    AskHoliday = False
                    SkipHoliday = False
                    If MsgBox(Msg$, 36) = 6 Then
                        SkipHoliday = True
                        GoTo nextAttendance
                    End If
                End If
            Else
                If SkipHoliday Then
                    GoTo nextAttendance
                End If
            End If
            
            rsAttD.AddNew
            rsAttD("AD_COMPNO") = "001"
            rsAttD("AD_EMPNBR") = rsempt("ED_EMPNBR")
            rsAttD("AD_DOA") = xDATE
            rsAttD("AD_REASON") = clpCode(4).Text
            rsAttD("AD_HRS") = fblDoul
            rsAttD("AD_COMM") = memComments
            If Len(Trim(txtShift)) = 0 Then
                rsAttD("AD_SHIFT") = GetJHData(rsempt("ED_EMPNBR"), "JH_SHIFT", "") 'txtShift   'Ticket #15323
            Else
                rsAttD("AD_SHIFT") = txtShift
            End If
            rsAttD("AD_SUPER") = IIf(Len(elpSup.Text) > 0, getEmpnbr(elpSup.Text), Null)
            rsAttD("AD_SEN") = IIf(chkSeniority.Value, 1, 0)
            rsAttD("AD_EMELEA") = IIf(chkEMELEA.Value, 1, 0)
            rsAttD("AD_INDICATOR") = IIf(chkIncentive.Value, 1, 0)
            rsAttD("AD_INCID") = IIf(chkIncident.Value, 1, 0)
            If IsNumeric(txtPoint) Then
                rsAttD("AD_POINT") = txtPoint
            End If
            rsAttD("AD_PAYROLL_ID") = rsempt("ED_PAYROLL_ID")
            rsAttD("AD_GLNO") = rsempt("ED_GLNO")
            rsAttD("AD_ORG") = rsempt("ED_ORG")
            
            If optEmployee(0) Then  'Release 8.0
                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsempt("ED_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            Else
                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM Term_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsempt("ED_EMPNBR") & " AND TERM_SEQ = " & rsAttD("TERM_SEQ")
                rsCurSal.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
            End If
            If Not rsCurSal.BOF Then
                If rsCurSal("SH_SALARY") > 0 Then
                    rsAttD("AD_SALARY") = rsCurSal("SH_SALARY")
                    rsAttD("AD_SALCD") = rsCurSal("SH_SALCD")
                End If
            End If
            rsCurSal.Close
                        
            If optEmployee(0) Then  'Release 8.0
                SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsempt("ED_EMPNBR")
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            Else
                SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM Term_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsempt("ED_EMPNBR") & " AND TERM_SEQ = " & rsAttD("TERM_SEQ")
                rsCurSal.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
            End If
            If Not rsCurSal.EOF Then
                rsAttD("AD_JOB") = rsCurSal("JH_JOB")
                rsAttD("AD_DHRS") = rsCurSal("JH_DHRS")
                rsAttD("AD_WHRS") = rsCurSal("JH_WHRS")
            End If
            rsCurSal.Close
            
            rsAttD("AD_LDATE") = Date
            rsAttD("AD_LTIME") = Time$
            rsAttD("AD_LUSER") = glbUserID
            
            'Ticket #18668 - 7.9 Enhancement
            rsAttD("AD_SOURCE") = "IHRATU"
                        
            rsAttD.Update
  
            '7.9 Enhancement
            If gsAttachment_DB Then
                If glbDocNewRecord Then 'New Record only
                    If Len(glbDocImpFile) > 0 Then
                        glbDocKey = rsAttD("AD_ATT_ID")
                        If optEmployee(0) Then 'Release 8.0
                            Call AttachmentAdd(rsempt("ED_EMPNBR"), glbDocImpFile, glbDocType, glbDocDesc)
                        Else
                            If glbtermopen = True Then
                                Call AttachmentAdd(rsempt("ED_EMPNBR"), glbDocImpFile, glbDocType, glbDocDesc)
                            Else
                                glbtermopen = True
                                Call AttachmentAdd(rsempt("ED_EMPNBR"), glbDocImpFile, glbDocType, glbDocDesc)
                                glbtermopen = False
                            End If
                        End If
                    End If
                End If
                'glbDocImpFile = ""
            End If
  
            'Ticket #12718
            xKey = rsempt("ED_EMPNBR")
            xKey = xKey & "|" & Format(xDATE, "dd-mmm-yyyy")
            xKey = xKey & "|" & clpCode(4).Text
            Call Attendance_Master_Integration(xKey, rsAttD("AD_ATT_ID"))
            
            If glbWFC Then 'Ticket #28919 Franks 03/24/2017
                Call WFC_Attend_To_AT(rsAttD("AD_EMPNBR"), "M", rsAttD("AD_DOA"), rsAttD("AD_REASON"), rsAttD("AD_ATT_ID"), , , "U")
            End If
        
nextAttendance:
            rsempt.MoveNext
        Loop
        rsempt.Close
        rsAttD.Close
'        'For regular clients
'        SQLQ = "INSERT INTO HR_ATTENDANCE "
'        SQLQ = SQLQ & "( AD_EMPNBR, AD_COMPNO, AD_DOA, AD_REASON, AD_HRS, AD_COMM, "
'        SQLQ = SQLQ & "AD_SHIFT, AD_SUPER, AD_INCID, AD_SEN, AD_EMELEA, AD_INDICATOR, AD_LDATE, AD_LTIME, AD_LUSER )  "
'
'        SQLQ = SQLQ & "SELECT ED_EMPNBR AS AD_EMPNBR, '001' AS AD_COMPNO, "
'        SQLQ = SQLQ & Date_SQL(xDATE) & " AS AD_DOA,"
'
'        SQLQ = SQLQ & "'" & clpCode(4).Text & "' AS AD_REASON, "  '
'        SQLQ = SQLQ & fblDoul & " AS AD_HRS,"
'        SQLQ = SQLQ & "'" & Replace(memComments, "'", "''") & "' AS AD_COMM,"
'        SQLQ = SQLQ & "'" & txtShift & "' AS AD_Shift,"
'        SQLQ = SQLQ & IIf(Len(elpSup.Text) > 0, getEmpnbr(elpSup.Text), "Null") & " AS AD_Super, "
'        SQLQ = SQLQ & IIf(chkIncident.Value, 1, 0) & " AS AD_IncID, "
'        SQLQ = SQLQ & IIf(chkSeniority.Value, 1, 0) & " AS AD_SEN, "
'        SQLQ = SQLQ & IIf(ChkEMELEA.Value, 1, 0) & " AS AD_EMELEA, "
'        SQLQ = SQLQ & IIf(chkIncentive.Value, 1, 0) & " AS AD_INDICATOR, " 'added by RAUBREY 6/3/97
'        SQLQ = SQLQ & Date_SQL(Date) & " AS AD_LDATE, "
'        SQLQ = SQLQ & "'" & Time$ & "' AS AD_LTIME, "
'        SQLQ = SQLQ & "'" & glbUserID & "' AS AD_LUser FROM HREMP WHERE" & fglbESQLQ
'        If xDup Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (" & TSQLQ & ")"
'        gdbAdoIhr001.Execute SQLQ
    
    Else
        If optEmployee(0) Then  'Release 8.0
            'For Whitby, get Incident and Disciplinary
            'Open a empty record of Attendance table
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_REASON ='**'"
            rsAttD.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE" & fglbESQLQ
            If xDup Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (" & TSQLQ & ")"
            rsempt.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsempt.EOF
                rsAttD.AddNew
                rsAttD("AD_COMPNO") = "001"
                rsAttD("AD_EMPNBR") = rsempt("ED_EMPNBR")
                rsAttD("AD_DOA") = xDATE
                rsAttD("AD_REASON") = clpCode(4).Text
                rsAttD("AD_HRS") = fblDoul
                rsAttD("AD_COMM") = memComments
                rsAttD("AD_SHIFT") = txtShift
                rsAttD("AD_SUPER") = IIf(Len(elpSup.Text) > 0, getEmpnbr(elpSup.Text), Null)
                rsAttD("AD_SEN") = IIf(chkSeniority.Value, 1, 0)
                rsAttD("AD_EMELEA") = IIf(chkEMELEA.Value, 1, 0)
                rsAttD("AD_INDICATOR") = IIf(chkIncentive.Value, 1, 0)
                If IsNumeric(txtPoint) Then
                    rsAttD("AD_POINT") = txtPoint
                End If
                'For Whitby only
                xDiscipFlag = False
                rsAttD("AD_INCID") = WhitbyGetIncidentFlags(rsempt("ED_EMPNBR"), CVDate(xDATE), clpCode(4).Text)
                rsAttD("AD_LDATE") = Date
                rsAttD("AD_LTIME") = Time$
                rsAttD("AD_LUSER") = glbUserID
                rsAttD.Update
                
                If xDiscipFlag Then
                    Call WhitbyUpdateDisciplinary(rsempt("ED_EMPNBR"), CVDate(xDATE), clpCode(4).Text)
                End If
                Call Whitby60daysRule(rsempt("ED_EMPNBR"), "")
                
                If glbWFC Then 'Ticket #28919 Franks 03/24/2017
                    Call WFC_Attend_To_AT(rsAttD("AD_EMPNBR"), "M", rsAttD("AD_DOA"), rsAttD("AD_REASON"), rsAttD("AD_ATT_ID"), , , "U")
                End If
            
                rsempt.MoveNext
            Loop
            rsempt.Close
            rsAttD.Close
        End If
    End If
    
'''    'Current Salary to AD_SALARY FOR Casey House
'''    'If glbCompSerial = "S/N - 2214W" Then jerry let make this for everone
'''    If Not glbVadim Then
'''        Dim xUpFlag As Boolean
'''        SQLQ = "SELECT AD_EMPNBR,AD_SALARY,AD_JOB,AD_ORG,AD_SALCD,AD_DHRS,AD_WHRS,AD_PAYROLL_ID,AD_GLNO FROM HR_ATTENDANCE "
'''        SQLQ = SQLQ & " WHERE AD_REASON = '" & clpCode(4).Text & "' "
'''        SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(xDATE)
'''        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
'''        rsAttSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'''        Do While Not rsAttSal.EOF
'''            xUpFlag = False
'''            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsAttSal("AD_EMPNBR")
'''            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
'''            If rsCurSal("SH_SALARY") > 0 Then
'''                rsAttSal("AD_SALARY") = rsCurSal("SH_SALARY")
'''                rsAttSal("AD_SALCD") = rsCurSal("SH_SALCD")
'''                xUpFlag = True
'''            End If
'''            rsCurSal.Close
'''
'''            SQLQ = "SELECT ED_EMPNBR, ED_ORG,ED_PAYROLL_ID,ED_GLNO FROM HREMP WHERE ED_EMPNBR = " & rsAttSal("AD_EMPNBR")
'''            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
'''            If Not rsCurSal.EOF Then
'''                rsAttSal("AD_ORG") = rsCurSal("ED_ORG")
'''                rsAttSal("AD_PAYROLL_ID") = rsCurSal("ED_PAYROLL_ID")
'''                rsAttSal("AD_GLNO") = rsCurSal("ED_GLNO")
'''                xUpFlag = True
'''            End If
'''            rsCurSal.Close
'''
'''            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsAttSal("AD_EMPNBR")
'''            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
'''            If Not rsCurSal.EOF Then
'''                rsAttSal("AD_JOB") = rsCurSal("JH_JOB")
'''                rsAttSal("AD_DHRS") = rsCurSal("JH_DHRS")
'''                rsAttSal("AD_WHRS") = rsCurSal("JH_WHRS")
'''                xUpFlag = True
'''            End If
'''            rsCurSal.Close
'''            If xUpFlag Then
'''                rsAttSal.Update
'''            End If
'''            rsAttSal.MoveNext
'''        Loop
'''        rsAttSal.Close
'''    End If

    xDATE = DateAdd("d", 1, xDATE)
Next

If optEmployee(0) Then  'Release 8.0
    If glbWFC And glbPlantCode = "WHBY" Then
        If yDiscipFlag Then
            Call cmdViewDiscip_Click
        End If
    End If
    If glbBurlTech Then 'BTI Points Recalculate
        Call BTIPoint(fglbESQLQ)
    End If
End If

Call EntReCalc(fglbESQLQ)
Call EntReCalcHr

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    Call ReCalcOvt("")
'End If

'Town of Ajax
If glbCompSerial = "S/N - 2173W" Then
    Call Recalculate_OTBANK
End If

If optEmployee(0) Then  'Release 8.0
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ
    If xDup Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (" & TSQLQ & ")"
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
Else
    SQLQ = "SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ
    If xDup Then SQLQ = SQLQ & " AND ED_EMPNBR NOT IN (" & TSQLQ & ")"
    RSEMPLIST.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly, adLockReadOnly
End If
Do While Not RSEMPLIST.EOF
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & RSEMPLIST("ED_EMPNBR")
    Else
        strEMPLIST = strEMPLIST & RSEMPLIST("ED_EMPNBR")
    End If
    RSEMPLIST.MoveNext
Loop
RSEMPLIST.Close

modInsRecs = True

Exit Function

modInsRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modInsRecs", "Attendance", "Insert")
modInsRecs = False
Resume Next

End Function

Private Function modUpdRecs()

Dim SQLQ As String
Dim rsempt As New ADODB.Recordset
Dim rsCurSal As New ADODB.Recordset
Dim rsAttD As New ADODB.Recordset
Dim fblDoul As Double
Dim DgDef, Title$, Msg$, Response%
Dim xExclude As Boolean
Dim xID

modUpdRecs = False
On Error GoTo modUpdRecs2_Err

Screen.MousePointer = HOURGLASS

xExclude = False

Call getWSQLQ
    'modified by Bryan Ticket# 11702 Sep 19, 2006
    
fblDoul = CDbl(medHours)
If UCase(clpCode(4).Text) = "OT15" Then fblDoul = fblDoul * 1.5
If UCase(clpCode(4).Text) = "OT20" Then fblDoul = fblDoul * 2

'City of Timmins - Ticket #16168
If glbCompSerial = "S/N - 2375W" Then
    If UCase(clpCode(4).Text) = "OT05" Then fblDoul = fblDoul * 0.5
    If UCase(clpCode(4).Text) = "OT25" Then fblDoul = fblDoul * 2.5
End If

If optEmployee(0) Then  'Release 8.0
    SQLQ = "SELECT * FROM HR_ATTENDANCE  "
Else
    SQLQ = "SELECT * FROM Term_ATTENDANCE  "
End If
SQLQ = SQLQ & " WHERE AD_REASON = '" & clpCode(4).Text & "' "
If Len(dlpToDate.Text) > 0 Then
    SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpAttDate.Text)
    SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpToDate.Text)
Else
    SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(dlpAttDate.Text)
End If

If optEmployee(0) Then  'Release 8.0
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
            
    'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted - Warn the users
    'Check if there are any ESS Approved records in this selection criteria
    If ContainsESSApprovedAttendance(SQLQ) Then
        Msg$ = "There are some ESS Approved Attendance records in this selection." & vbCrLf & vbCrLf & "Do you want exclude them from Update?"
        Response% = MsgBox(Msg$, vbExclamation + vbYesNoCancel, "ESS Approved Attendance Records found")      ' Get user response.
        If Response = IDCANCEL Then
            Screen.MousePointer = DEFAULT
            Exit Function
        ElseIf Response = IDYES Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
            xExclude = True
        End If
    End If
    
    rsAttD.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdText
Else
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & ")"
    
    'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted - Warn the users
    'Check if there are any ESS Approved records in this selection criteria
    If ContainsESSApprovedAttendance(SQLQ) Then
        Msg$ = "There are some ESS Approved Attendance records in this selection." & vbCrLf & vbCrLf & "Do you want exclude them from Update?"
        Response% = MsgBox(Msg$, vbExclamation + vbYesNoCancel, "ESS Approved Attendance Records found")    ' Get user response.
        If Response = IDCANCEL Then
            Screen.MousePointer = DEFAULT
            Exit Function
        ElseIf Response = IDYES Then
            SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
            xExclude = True
        End If
    End If
    
    rsAttD.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdText
End If
If rsAttD.EOF = False And rsAttD.BOF = False Then
    Do
        rsAttD("AD_COMPNO") = "001"
        'rsAttD("AD_DOA") = dlpAttDate.Text
        rsAttD("AD_REASON") = clpCode(4).Text
        rsAttD("AD_HRS") = fblDoul
        rsAttD("AD_COMM") = memComments
        rsAttD("AD_SHIFT") = txtShift
        rsAttD("AD_SUPER") = IIf(Len(elpSup.Text) > 0, getEmpnbr(elpSup.Text), Null)
        rsAttD("AD_SEN") = IIf(chkSeniority.Value, 1, 0)
        rsAttD("AD_EMELEA") = IIf(chkEMELEA.Value, 1, 0)
        rsAttD("AD_INDICATOR") = IIf(chkIncentive.Value, 1, 0)
        rsAttD("AD_INCID") = IIf(chkIncident.Value, 1, 0)
        If IsNumeric(txtPoint) Then
            rsAttD("AD_POINT") = txtPoint
        End If
        
        If optEmployee(0) Then  'Release 8.0
            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO FROM HREMP WHERE ED_EMPNBR=" & rsAttD("AD_EMPNBR")
            rsempt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Else
            SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_ORG,ED_GLNO FROM Term_HREMP WHERE ED_EMPNBR=" & rsAttD("AD_EMPNBR") & " AND TERM_SEQ = " & rsAttD("TERM_SEQ")
            rsempt.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        End If
        If rsAttD.EOF = False And rsAttD.BOF = False Then
            rsAttD("AD_PAYROLL_ID") = rsempt("ED_PAYROLL_ID")
            rsAttD("AD_GLNO") = rsempt("ED_GLNO")
            rsAttD("AD_ORG") = rsempt("ED_ORG")
        End If
        rsempt.Close
        
        If optEmployee(0) Then  'Release 8.0
            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsAttD("AD_EMPNBR")
            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM Term_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & rsAttD("AD_EMPNBR") & " AND TERM_SEQ = " & rsAttD("TERM_SEQ")
            rsCurSal.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
        If Not rsCurSal.BOF Then
            If rsCurSal("SH_SALARY") > 0 Then
                rsAttD("AD_SALARY") = rsCurSal("SH_SALARY")
                rsAttD("AD_SALCD") = rsCurSal("SH_SALCD")
            End If
        End If
        rsCurSal.Close
        
        If optEmployee(0) Then  'Release 8.0
            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsAttD("AD_EMPNBR")
            rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        Else
            SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB,JH_DHRS,JH_WHRS FROM Term_JOB_HISTORY WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & rsAttD("AD_EMPNBR") & " AND TERM_SEQ = " & rsAttD("TERM_SEQ")
            rsCurSal.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        End If
        If Not rsCurSal.EOF Then
            rsAttD("AD_JOB") = rsCurSal("JH_JOB")
            rsAttD("AD_DHRS") = rsCurSal("JH_DHRS")
            rsAttD("AD_WHRS") = rsCurSal("JH_WHRS")
        End If
        rsCurSal.Close
        
        rsAttD("AD_LDATE") = Date
        rsAttD("AD_LTIME") = Time$
        rsAttD("AD_LUSER") = glbUserID
        
        'Ticket #18668 - 7.9 Enhancement
        rsAttD("AD_SOURCE") = "IHRATU"

        rsAttD.Update
        
        'Ticket #12718
        xKey = rsAttD("AD_EMPNBR")
        xKey = xKey & "|" & Format(rsAttD("AD_DOA"), "dd-mmm-yyyy")
        xKey = xKey & "|" & clpCode(4).Text
        Call Attendance_Master_Integration(xKey, rsAttD("AD_ATT_ID"))
        
        If glbWFC Then 'Ticket #28919 Franks 03/24/2017
            Call WFC_Attend_To_AT(rsAttD("AD_EMPNBR"), "M", rsAttD("AD_DOA"), rsAttD("AD_REASON"), rsAttD("AD_ATT_ID"), , , "U")
        End If

        rsAttD.MoveNext
    Loop Until rsAttD.EOF
End If
rsAttD.Close
'end bryan

If glbBurlTech Then 'BTI Points Recalculate
    Call BTIPoint(fglbESQLQ)
End If
Call EntReCalc(fglbESQLQ)
Call EntReCalcHr

'Town of Aurora
'If glbCompSerial = "S/N - 2378W" Then
    Call ReCalcOvt("")
'End If

'Town of Ajax
If glbCompSerial = "S/N - 2173W" Then
    Call Recalculate_OTBANK
End If

'Call VacSickHourlyFollowUp(clpCode(4).Text, dlpAttDate)

If optEmployee(0) Then  'Release 8.0
    SQLQ = "SELECT distinct AD_EMPNBR FROM HR_ATTENDANCE WHERE AD_REASON = '" & clpCode(4).Text & "' "
Else
    SQLQ = "SELECT distinct AD_EMPNBR FROM Term_ATTENDANCE WHERE AD_REASON = '" & clpCode(4).Text & "' "
End If
If Len(dlpToDate.Text) > 0 Then
    SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpAttDate.Text)
    SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpToDate.Text)
Else
    SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(dlpAttDate.Text)
End If
If optEmployee(0) Then  'Release 8.0
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
    
    'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted
    If xExclude Then
        SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
    End If
    
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
Else
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & fglbESQLQ & ")"
    
    'Ticket #25268: ESS Approved Attendance record cannot be editted or deleted
    If xExclude Then
        SQLQ = SQLQ & " AND (AD_SOURCE <> 'ESSAP' OR AD_REQID IS NULL)"
    End If
    
    RSEMPLIST.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly, adLockReadOnly
End If
Do While Not RSEMPLIST.EOF
    If Len(strEMPLIST) > 0 Then
        strEMPLIST = strEMPLIST & "," & RSEMPLIST("AD_EMPNBR")
    Else
        strEMPLIST = strEMPLIST & RSEMPLIST("AD_EMPNBR")
    End If
    RSEMPLIST.MoveNext
Loop
RSEMPLIST.Close

modUpdRecs = True
glbflgFU = False

Set rsAttD = Nothing
Set rsempt = Nothing
Set rsCurSal = Nothing

Exit Function

modUpdRecs2_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdRecs", "Attendance Reason", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub optEmployee_Click(Index As Integer)
    If optEmployee(1) Then
        MsgBox "Only 'Mass Delete' option is allowed for 'Terminated Employees'", vbInformation
        elpEEID.LookupType = TERM
        
        lblTitle(1).Visible = True
        dlpDateRange(0).Visible = True
        dlpDateRange(1).Visible = True
    Else
        elpEEID.LookupType = 0
        
        lblTitle(1).Visible = False
        dlpDateRange(0).Visible = False
        dlpDateRange(1).Visible = False
    End If
End Sub

Private Sub scrControl_Change()
pcSelectCri.Top = 120 - scrControl.Value
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub getWSQLQ()

fglbESQLQ = glbSeleDeptUn

'Release 8.0 - Multiple selections
If Len(clpDept.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO IN ('" & Replace(clpDept.Text, ",", "','") & "') "
If Len(clpDiv.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV IN ('" & Replace(clpDiv.Text, ",", "','") & "') "
If Len(clpCode(2).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
If Len(clpCode(3).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
If Len(clpCode(0).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
If Len(clpCode(1).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(1).Text, ",", "','") & "') "
If Len(clpCode(6).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "

If Len(clpPT.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "') "

If glbLinamar Then
    If Len(clpCode(5).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND (ED_REGION = '" & clpDiv.Text & clpCode(5).Text & "' or  ED_REGION= 'ALL" & clpCode(5).Text & "')"
Else
    If Len(clpCode(5).Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_REGION IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
End If

If Len(elpEEID.Text) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
'Frank 02/10/04 Ticket #5522
'Disable the use of Attendance Mass Update for Casey House
If glbCompSerial = "S/N - 2214W" Then
    UpdateRight = False
Else
    UpdateRight = GetMassUpdateSecurities("Attendance_MassUpdate", glbUserID) 'gSec_Upd_Attendance
End If
End Property

Public Property Get Addable() As Boolean
    'Ticket #7500 - Town of Ajax
    'Ticket #10423 Jerry asked to removed serial#control
  '  If glbCompSerial = "S/N - 2173W" Then
        Addable = gSec_Add_Attendance
   ' Else
    '    Addable = True
    'End If
End Property

Public Property Get Updateble() As Boolean
'Updateble = True
'Ticket #7500 - Town of Ajax
'Ticket #10423 Jerry asked to removed serial#control 05/09/2006
'If glbCompSerial = "S/N - 2173W" Then
    Updateble = gSec_Upd_Attendance
'End If

End Property

Public Property Get Deleteble() As Boolean
'Deleteble = True
'Ticket #7500 - Town of Ajax
'Ticket #10423 Jerry asked to removed serial#control
'If glbCompSerial = "S/N - 2173W" Then
    Deleteble = gSec_Upd_Attendance
'End If

End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Function WhitbyGetIncidentFlags(xEmpNo, xDOA, xReason)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xVPoint, xNextDiscipStep
Dim xCodeFlag As Boolean
Dim xIncidentAmt, xTmpDate1, xTmpDate2
Dim xIncidentVal, xDayAmt, I, xDayDiff
    'Check Attendance Code if it has Absend checked and has Points
    xCodeFlag = True
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = '" & xReason & "' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF Then 'No this code
        xCodeFlag = False
    Else
        'If Not rsTemp("TB_ABSENCE") Then xCodeFlag = False 'Absence was unchecked
        If IsNull(rsTemp("TB_USR2")) Then xCodeFlag = False 'Point is null
        If rsTemp("TB_USR2") = 0 Then xCodeFlag = False 'Point is 0
    End If
    If IsNull(rsTemp("TB_USR2")) Then
        xVPoint = 0
    Else
        xVPoint = rsTemp("TB_USR2")
    End If
    rsTemp.Close
    If Not xCodeFlag Then
        'rsDATA("AD_INCID") = False
        WhitbyGetIncidentFlags = 0
        Exit Function
    End If
    
    'If xNextDiscipStep = 0  then Check the first three occurences, else go to next Disciplinary Step
    SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_REASON,AD_INCID,TB_ABSENCE,TB_USR2 FROM HR_ATTENDANCE "
    SQLQ = SQLQ & "LEFT JOIN HRTABL ON (HR_ATTENDANCE.AD_REASON = HRTABL.TB_KEY) AND (HR_ATTENDANCE.AD_REASON_TABL = HRTABL.TB_NAME) "
    SQLQ = SQLQ & "WHERE AD_EMPNBR =" & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    SQLQ = SQLQ & "ORDER BY AD_DOA "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xIncidentAmt = 0
    xTmpDate1 = CVDate(glbDiscipStartDate)
    xTmpDate2 = CVDate(glbDiscipStartDate)
    Do While Not rsTemp.EOF
        If rsTemp("AD_INCID") Then
            xIncidentAmt = xIncidentAmt + 1
        End If
        'If rsTemp("TB_ABSENCE") Then
            If Not IsNull(rsTemp("TB_USR2")) Then
                If rsTemp("TB_USR2") > 0 Then
                    If rsTemp("AD_REASON") = xReason Then 'Check the same reason code
                        xTmpDate1 = rsTemp("AD_DOA")
                    End If
                End If
            End If
        'End If
        xTmpDate2 = rsTemp("AD_DOA")
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    'xIncidentVal
    'If xTmpDate1 = xTmpDate2 And (xTmpDate1 <> CVDate(glbDiscipStartDate)) Then
    '    'it means the last Attendance record was Absent,
    '    'don't turn on the Absent flag for the new record
    '    xIncidentVal = False
    'Else
    '    xIncidentVal = True
    '    xIncidentAmt = xIncidentAmt + 1
    'End If
    xDayAmt = DateDiff("d", xTmpDate1, xDOA): xDayDiff = 0
    If xDayAmt > 10 Then
        xDayDiff = 10
    Else
        For I = 1 To xDayAmt
            xTmpDate1 = DateAdd("d", 1, xTmpDate1)
            If Not (Weekday(xTmpDate1) = 1 Or Weekday(xTmpDate1) = 7) Then
                xDayDiff = xDayDiff + 1
            End If
        Next I
    End If
    If xDayDiff > 1 Then
        xIncidentVal = True
        xIncidentAmt = xIncidentAmt + 1
    Else
        xIncidentVal = False
    End If
    'rsDATA("AD_INCID") = xIncidentVal
    xOccuAmount = xIncidentAmt
    WhitbyGetIncidentFlags = xIncidentVal 'xIncidentAmt

    xDiscipFlag = True

End Function

Private Sub WhitbyUpdateDisciplinary(xEmpNo, xDOA, xReason) ', xHrs, xFlag)
Dim rsTemp As New ADODB.Recordset
Dim rsTem2 As New ADODB.Recordset
Dim SQLQ, xVPoint, xNextDiscipStep, xNextStepPlus
Dim xCodeFlag As Boolean
Dim CurDiscip, NextDiscip, xREPTAU1

    ''Disable it until Whitby is ready
    'Exit Sub
    
    'Check what is the Next Disciplinary Step
    xNextDiscipStep = 1
    SQLQ = "SELECT ED_EMPNBR,ED_DISCIPLINENEXT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("ED_DISCIPLINENEXT")) Then
            xNextDiscipStep = rsTemp("ED_DISCIPLINENEXT")
        End If
    End If
    rsTemp.Close
    
    If xNextDiscipStep = 1 And xOccuAmount <= 3 Then
        'if less than 3 Incident occurences, no Disciplianry action
        Exit Sub
    End If
    
    'Find the Disciplinary Code
    CurDiscip = "***": NextDiscip = "***"
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS ORDER BY DS_STEPNO " ' WHERE DS_STEPNO = " & xNextDiscipStep
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Sub
    Else
        Do While Not rsTemp.EOF
            If rsTemp("DS_STEPNO") = xNextDiscipStep - 1 Then
                CurDiscip = rsTemp("DS_DISCIPLINE")
            End If
            If rsTemp("DS_STEPNO") = xNextDiscipStep Then
                NextDiscip = rsTemp("DS_DISCIPLINE")
            End If
            rsTemp.MoveNext
        Loop
    End If
    rsTemp.Close
    
    'If Next Disciplinary action doesn't exist, exit sub
    If NextDiscip = "***" Then
        Exit Sub
    End If
    'Check if the Current Disciplianry Action exists
    '   If it doesn't exist, create a new one using next Step
    '   If it exists, check if the Counselling Date has beed entered
    '       if not entered, exit sub, don't do anything
    '       if entered, it means the Disciplinary Action was done by HR person, create a new action
    SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_TYPE = '" & CurDiscip & "' "
    SQLQ = SQLQ & "AND CL_LDATE >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsDate(rsTemp("CL_COUDATE")) Then
            'if not entered, exit sub, don't do anything
            rsTemp.Close
            Exit Sub
        End If
    End If
    rsTemp.Close
    
    'Reset the current Disciplinary to False before creating a new current
    SQLQ = "UPDATE HR_COUNSEL SET CL_COMPLETED = 0 WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_COMPLETED <> 0 "
    gdbAdoIhr001.Execute SQLQ
    
    xNextStepPlus = xNextDiscipStep + 1
    'Create Next Disciplinary Action
    SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_TYPE = '" & NextDiscip & "' "
    SQLQ = SQLQ & "AND CL_LDATE >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    SQLQ = SQLQ & "AND CL_REASON = 'ATT' "
    SQLQ = SQLQ & "AND CL_INCDATE= " & Date_SQL(xDOA) & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTemp.EOF Then
        
        'Get Next Step Number
        
        'Get Report #1 from current position
        xREPTAU1 = ""
        SQLQ = "SELECT JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & xEmpNo & " "
        If rsTem2.State <> 0 Then rsTem2.Close
        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTem2.EOF Then
            If Not IsNull(rsTem2("JH_REPTAU")) Then
                xREPTAU1 = Trim(rsTem2("JH_REPTAU"))
            End If
        End If
        rsTem2.Close
        
        ''Put Disciplinary Action in Attendance table
        'SQLQ = "UPDATE HR_ATTENDANCE SET AD_DISCIPLINE = '" & NextDiscip & "' "
        'SQLQ = SQLQ & "WHERE AD_EMPNBR = " & xEmpNo & " "
        'SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xDOA) & " "
        'SQLQ = SQLQ & "AND AD_REASON = '" & xReason & "' "
        'gdbAdoIhr001.Execute SQLQ
        
        rsTemp.AddNew
        rsTemp("CL_COMPNO") = "001"
        rsTemp("CL_EMPNBR") = xEmpNo
        rsTemp("CL_INCDATE") = xDOA
        rsTemp("CL_TYPE") = NextDiscip
        If Len(xREPTAU1) > 0 Then rsTemp("CL_COUBY") = xREPTAU1
        rsTemp("CL_LDATE") = Date
        rsTemp("CL_LTIME") = Time$
        rsTemp("CL_LUSER") = glbUserID
        rsTemp("CL_ATTDATE") = xDOA
        rsTemp("CL_ATTREASON") = xReason
        rsTemp("CL_COMPLETED") = -1
    Else
        rsTemp("CL_COUDATE") = Null
        rsTemp("CL_INCDATE") = xDOA
        rsTemp("CL_LDATE") = Date
        rsTemp("CL_LTIME") = Time$
        rsTemp("CL_LUSER") = glbUserID
        rsTemp("CL_ATTDATE") = xDOA
        rsTemp("CL_ATTREASON") = xReason
        rsTemp("CL_COMPLETED") = -1
    End If
    rsTemp.Update
    
    'Put Disciplinary Action in Attendance table
    SQLQ = "UPDATE HR_ATTENDANCE SET AD_DISCIPLINE = '" & NextDiscip & "' "
    SQLQ = SQLQ & "WHERE AD_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xDOA) & " "
    SQLQ = SQLQ & "AND AD_REASON = '" & xReason & "' "
    gdbAdoIhr001.Execute SQLQ
        
    'Create a report records
    SQLQ = "DELETE FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "SELECT * FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
    If rsTem2.State <> 0 Then rsTem2.Close
    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsTem2.AddNew
    rsTem2("AD_COMPNO") = "001"
    rsTem2("AD_EMPNBR") = xEmpNo
    rsTem2("AD_DOA") = xDOA
    rsTem2("AD_REASON") = xReason
    rsTem2("AD_DISCIPLINE") = NextDiscip
    'rsTem2("AD_POINT") = ""
    rsTem2("AD_LDATE") = Date
    rsTem2("AD_LTIME") = Time$
    rsTem2("AD_WRKEMP") = glbUserID
    rsTem2.Update
    rsTem2.Close
    yDiscipFlag = True
    'Call cmdViewDiscip_Click
    rsTemp.Close
    
    If xNextStepPlus > xNextDiscipStep Then
        SQLQ = "UPDATE HREMP SET ED_DISCIPLINENEXT = " & xNextStepPlus & " "
        SQLQ = SQLQ & "WHERE ED_EMPNBR = " & xEmpNo
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Private Sub cmdViewDiscip_Click()
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzdiscip.rpt"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.SubreportToChange = "AttDetail"
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.SubreportToChange = ""
    Me.vbxCrystal.SelectionFormula = "{HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.WindowTitle = "Disciplinary Report"
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
End Sub

Private Function getWSQLQRPT() As String
'getWSQLQRPT = glbSeleDeptUn    'Department security removed by Bryan, redundant, this is a list of changes, whether they have security is irrelevant at this point
'If Len(clpDept.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DEPTNO} = '" & clpDept.Text & "')"
'If Len(clpDiv.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_DIV} = '" & clpDiv.Text & "') "
'If Len(clpCode(1).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_LOC} = '" & clpCode(1).Text & "') "
'If Len(clpCode(2).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ORG} = '" & clpCode(2).Text & "') "
'If Len(clpCode(3).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_EMP} = '" & clpCode(3).Text & "') "
'If Len(clpCode(5).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_REGION} = '" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(5).Text & "') "
'If Len(clpCode(6).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_ADMINBY} = '" & clpCode(6).Text & "') "
'If Len(clpCode(7).Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_BENEFIT_GROUP} = '" & clpCode(7).Text & "') "
'If Len(clpPT.Text) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.ED_PT} = '" & clpPT.Text & "') "
If Len(strEMPLIST) > 0 Then getWSQLQRPT = " ({HREMP.ED_EMPNBR} IN [" & strEMPLIST & "]) "

If optEmployee(1) Then  'Release 8.0
    If Len(strTERMSEQ) > 0 Then getWSQLQRPT = getWSQLQRPT & " AND ({HREMP.TERM_SEQ} IN [" & strTERMSEQ & "]) "
End If

End Function

Private Sub Recalculate_OTBANK()
Dim rsEmp As New ADODB.Recordset
Dim rsAttend As New ADODB.Recordset
Dim rsAttendCT As New ADODB.Recordset
Dim SQLQ

'Set ED_OTBANK to zero for the first time otherwise Null will be updated if some Value - Null
SQLQ = "UPDATE HREMP SET ED_OTBANK = 0"
gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT ED_EMPNBR, ED_OTBANK FROM HREMP"
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

If Not rsEmp.EOF Then
    rsEmp.MoveFirst
    
    Do While Not rsEmp.EOF
        
        If glbOracle Then
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE substr(AD_REASON,1,2) = 'OT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE substr(AD_REASON,1,2) = 'CT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        Else
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'OT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'CT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        End If
        If Not rsAttend.EOF Then
            If Not rsAttendCT.EOF Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        Else
            If Not rsAttendCT.EOF Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK = " & 0 - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
        rsAttend.Close
        rsAttendCT.Close
        
        rsEmp.MoveNext
    Loop
End If
rsEmp.Close

End Sub

Private Function getRecordCount_Update()
    Dim SQLQ As String
    Dim rsAttD As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Update = 0
    recCount = 0

    Call getWSQLQ

    If optEmployee(0) Then  'Release 8.0
        SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOT_REC FROM HR_ATTENDANCE  "
    Else
        SQLQ = "SELECT COUNT(AD_EMPNBR) AS TOT_REC FROM TERM_ATTENDANCE  "
    End If
    SQLQ = SQLQ & " WHERE AD_REASON = '" & clpCode(4).Text & "' "
    If Len(dlpToDate.Text) > 0 Then
        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpAttDate.Text)
        SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpToDate.Text)
    Else
        SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(dlpAttDate.Text)
    End If
    If optEmployee(0) Then  'Release 8.0
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
        rsAttD.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM TERM_HREMP WHERE " & fglbESQLQ & ")"
        rsAttD.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    End If
    If Not rsAttD.EOF Then
        recCount = rsAttD("TOT_REC")
    Else
        recCount = 0
    End If
    rsAttD.Close
    Set rsAttD = Nothing
    
    getRecordCount_Update = recCount

End Function

Private Function getRecordCount_Delete()
    Dim SQLQ As String
    Dim rsAttD As New ADODB.Recordset
    Dim recCount As Integer
    Dim WSQLQ As String
    
    getRecordCount_Delete = 0
    recCount = 0

    Call getWSQLQ
    
    WSQLQ = WSQLQ & " AD_REASON = '" & clpCode(4).Text & "' "
    
    If Len(dlpToDate.Text) > 0 Then
        If Len(dlpAttDate.Text) > 0 Then
            WSQLQ = WSQLQ & " AND AD_DOA >= " & Date_SQL(dlpAttDate.Text)
        End If
        WSQLQ = WSQLQ & " AND AD_DOA <= " & Date_SQL(dlpToDate.Text)
    Else
        If Len(dlpAttDate.Text) > 0 Then
            WSQLQ = WSQLQ & " AND AD_DOA = " & Date_SQL(dlpAttDate.Text)
        End If
    End If
    
    If Len(medHours.Text) > 0 Then
        WSQLQ = WSQLQ & " AND AD_HRS = " & medHours.Text
    End If
    
    If glbCompSerial = "S/N - 2192W" Then   'County of Essex
        'Exclude Machine # populated records
        If Len(medHours.Text) > 0 Then 'Ticket #12338
            If medHours.Text = 0 Then
                WSQLQ = WSQLQ & " AND (AD_MACHINE_NUM IS NULL OR AD_MACHINE_NUM = '')"
            End If
        End If
    End If
    
    If optEmployee(0) Then  'Release 8.0
        SQLQ = "SELECT COUNT(AD_ATT_ID) AS TOT_REC FROM HR_ATTENDANCE WHERE " & WSQLQ
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & fglbESQLQ & ")"
        rsAttD.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT COUNT(AD_ATT_ID) AS TOT_REC FROM TERM_ATTENDANCE WHERE " & WSQLQ
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM TERM_HREMP WHERE " & fglbESQLQ & ")"
        
        'Termination Date Range
        If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
            SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
            'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
            SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0))
            SQLQ = SQLQ & " AND Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
        Else
            If IsDate(dlpDateRange(0)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT >=" & Date_SQL(dlpDateRange(0)) & ")"
            End If
            If IsDate(dlpDateRange(1)) Then
                SQLQ = SQLQ & " AND TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HRTRMEMP "
                'SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
                SQLQ = SQLQ & " WHERE Term_DOT <=" & Date_SQL(dlpDateRange(1)) & ")"
            End If
        End If
        
        rsAttD.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    End If
        
    If Not rsAttD.EOF Then
        recCount = rsAttD("TOT_REC")
    Else
        recCount = 0
    End If
    rsAttD.Close
    Set rsAttD = Nothing
    
    getRecordCount_Delete = recCount

End Function

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    Call getWSQLQ
    
    If optEmployee(0) Then  'Release 8.0
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP WHERE " & fglbESQLQ
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM TERM_HREMP WHERE " & fglbESQLQ
        rsEmp.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    End If
    If Not rsEmp.EOF Then
        recCount = rsEmp("TOT_REC")
    Else
        recCount = 0
    End If
    rsEmp.Close
    Set rsEmp = Nothing
    
    getRecordCount_Add = recCount

End Function

Private Function ContainsESSApprovedAttendance(xSQLQ)
    Dim rsAttend As New ADODB.Recordset
        
    ContainsESSApprovedAttendance = False
    
    rsAttend.Open xSQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Do While Not rsAttend.EOF
        If rsAttend("AD_SOURCE") = "ESSAP" And Not IsNull(rsAttend("AD_REQID")) Then
            ContainsESSApprovedAttendance = True
            Exit Do
        Else
            ContainsESSApprovedAttendance = False
        End If
        
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing

End Function
