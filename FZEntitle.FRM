VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmREntitle 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Entitlements Report"
   ClientHeight    =   10260
   ClientLeft      =   255
   ClientTop       =   615
   ClientWidth     =   14325
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
   ScaleHeight     =   10260
   ScaleMode       =   0  'User
   ScaleWidth      =   14325
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   8235
      LargeChange     =   315
      Left            =   9720
      Max             =   100
      SmallChange     =   315
      TabIndex        =   55
      Top             =   600
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   54
      Top             =   8955
      Width           =   9735
   End
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   9495
      Begin VB.ComboBox cmbHours 
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
         ItemData        =   "FZEntitle.frx":0000
         Left            =   2220
         List            =   "FZEntitle.frx":0002
         TabIndex        =   8
         Tag             =   "Choose display in hours or days"
         Text            =   "Combo1"
         Top             =   2940
         Width           =   1215
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
         Index           =   0
         Left            =   2005
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Tag             =   "First Level of grouping records"
         Top             =   8085
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
         Left            =   2005
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Tag             =   "First Level of grouping records"
         Top             =   8400
         Width           =   2325
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   12
         Tag             =   "00-Employee Position Shift"
         Top             =   4290
         Visible         =   0   'False
         Width           =   450
      End
      Begin Threed.SSCheck chkAtt 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Tag             =   "If X-Show Attendance Details"
         Top             =   6315
         Width           =   2235
         _Version        =   65536
         _ExtentX        =   3942
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Attendance Details"
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
      Begin Threed.SSCheck chkEntitle 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Tag             =   "If X-Show Hourly Entitlements"
         Top             =   5730
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Hourly Entitlements"
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
      Begin Threed.SSCheck chkCompTm 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Tag             =   "If X-Show Compensatory Time"
         Top             =   6030
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Compensatory Time"
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   4
         Tag             =   "00-Enter Status Code"
         Top             =   1590
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDEM"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   1920
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TABLName        =   "EDPT"
         MaxLength       =   0
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1260
         Width           =   7395
         _ExtentX        =   13044
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
         Top             =   930
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
         Top             =   600
         Width           =   7395
         _ExtentX        =   13044
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
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   270
         Width           =   7395
         _ExtentX        =   13044
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
         TabIndex        =   11
         Tag             =   "00-Enter Section Code"
         Top             =   3960
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
         TabIndex        =   10
         Tag             =   "00-Enter Administered By Code"
         Top             =   3630
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
         TabIndex        =   9
         Tag             =   "00-Enter Region Code"
         Top             =   3300
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Tag             =   "10-Enter Employee Number"
         Top             =   2250
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         TextBoxWidth    =   7075
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3570
         TabIndex        =   15
         Tag             =   "40-Date upto and including this date forward"
         Top             =   4635
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   14
         Tag             =   "40-Date from and including this date forward"
         Top             =   4635
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin Threed.SSCheck chkSick 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Tag             =   "If X-Show Attendance Details"
         Top             =   6615
         Visible         =   0   'False
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Sick Time Only"
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   5280
         TabIndex        =   13
         Tag             =   "00-Enter Section Code"
         Top             =   4290
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   8
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   3
         Left            =   3570
         TabIndex        =   17
         Tag             =   "40-Compensatory Time Date upto and including this date forward"
         Top             =   4965
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   16
         Tag             =   "40-Compensatory Time Date from and including this date forward"
         Top             =   4965
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptau 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Tag             =   "10-Reporting Authority"
         Top             =   2595
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TextBoxWidth    =   7075
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkShowHrlyDetail 
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Tag             =   "If X-Show Attendance Details"
         Top             =   6315
         Visible         =   0   'False
         Width           =   1755
         _Version        =   65536
         _ExtentX        =   3096
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Hourly Entitlement"
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
      Begin Threed.SSCheck chkShowCompDetail 
         Height          =   255
         Left            =   6600
         TabIndex        =   25
         Tag             =   "If X-Show Attendance Details"
         Top             =   6315
         Visible         =   0   'False
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Compensatory Time"
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
      Begin Threed.SSCheck chkShowVacSickDetail 
         Height          =   255
         Left            =   2760
         TabIndex        =   23
         Tag             =   "If X-Show Attendance Details"
         Top             =   6315
         Visible         =   0   'False
         Width           =   1875
         _Version        =   65536
         _ExtentX        =   3307
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Vacation and Sick"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Value           =   -1  'True
      End
      Begin Threed.SSCheck chkShowCalculated 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Tag             =   "If X-Show Entitlement Cost"
         Top             =   7200
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Calculated"
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
      Begin Threed.SSCheck chkEntBalCost 
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Tag             =   "If X-Enter As of Date"
         Top             =   6900
         Visible         =   0   'False
         Width           =   3315
         _Version        =   65536
         _ExtentX        =   5847
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Entitlement Balance/Costs 'As of'"
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
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   4
         Left            =   6120
         TabIndex        =   29
         Tag             =   "40-Entitlement Balance/Cost As of Date"
         Top             =   6651
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
         Enabled         =   0   'False
      End
      Begin Threed.SSCheck chkCost 
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Tag             =   "If X-Show Entitlement Cost"
         Top             =   6900
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Entitlement Cost Only"
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
      Begin Threed.SSCheck chkShowSickOnly 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "If X-Show Attendance Details"
         Top             =   5430
         Visible         =   0   'False
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Sick Time Only"
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
      Begin Threed.SSCheck chkShowVacOnly 
         Height          =   255
         Left            =   2520
         TabIndex        =   19
         Tag             =   "If X-Show Attendance Details"
         Top             =   5430
         Visible         =   0   'False
         Width           =   2115
         _Version        =   65536
         _ExtentX        =   3731
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Vacation Only"
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
         Left            =   60
         TabIndex        =   53
         Top             =   315
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
         Left            =   60
         TabIndex        =   52
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lblTitle 
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
         Index           =   2
         Left            =   60
         TabIndex        =   51
         Top             =   2295
         Width           =   1290
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
         Left            =   60
         TabIndex        =   50
         Top             =   1305
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
         Left            =   60
         TabIndex        =   49
         Top             =   1635
         Width           =   450
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
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
         Left            =   60
         TabIndex        =   48
         Top             =   3000
         Width           =   510
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
      Begin VB.Label lblRepGrp 
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
         TabIndex        =   46
         Top             =   7755
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
         Left            =   90
         TabIndex        =   45
         Top             =   8115
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
         Index           =   1
         Left            =   90
         TabIndex        =   44
         Top             =   8430
         Width           =   660
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
         Left            =   60
         TabIndex        =   43
         Top             =   975
         Width           =   615
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
         Left            =   60
         TabIndex        =   42
         Top             =   3675
         Width           =   1125
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   60
         TabIndex        =   41
         Top             =   3345
         Width           =   510
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
         Left            =   60
         TabIndex        =   40
         Top             =   4005
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
         Left            =   60
         TabIndex        =   39
         Top             =   1965
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
         Left            =   60
         TabIndex        =   38
         Top             =   4335
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vacation Entitl. Date"
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
         Left            =   60
         TabIndex        =   37
         Top             =   4680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblSalDist 
         Caption         =   "Salary Distribution"
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
         Left            =   3660
         TabIndex        =   36
         Top             =   4335
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCompTimeDt 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comp. Time Date"
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
         Left            =   60
         TabIndex        =   35
         Top             =   5010
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblRepAuth 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Rept. Authority 1"
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
         Left            =   60
         TabIndex        =   34
         Top             =   2640
         Width           =   1185
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9480
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      WindowControls  =   -1  'True
      MarginTop       =   720
      MarginBottom    =   720
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmREntitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jaddy 10/26/99 add a CheckBox and edit lblTitle(5).caption="Display"

Option Explicit
 Dim medSalary


Private Function cri_cmpDates() As String
Dim retVal As String

    If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
        retVal = "Date Range: "
        If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
            retVal = retVal & dlpDateRange(2).Text & " TO " & dlpDateRange(3).Text
        ElseIf Len(dlpDateRange(2).Text) > 0 Then
            retVal = retVal & "FROM " & dlpDateRange(2).Text
        ElseIf Len(dlpDateRange(3).Text) > 0 Then
            retVal = retVal & "TO " & dlpDateRange(3).Text
        End If
        retVal = "lblDateRange='" & retVal & "'"
    Else
        retVal = "lblDateRange=''"
    End If
    cri_cmpDates = retVal

End Function

Private Sub chkAtt_Click(Value As Integer)
    If chkAtt Then
        chkCompTm = False: chkEntitle = False: chkCost = False 'Jaddy 11/5/99
        If chkSick.Visible Then chkSick = False
        
        'Show the Vacation, Sick, Hourly and Comp Time details checkbox
        chkShowHrlyDetail.Visible = True
        chkShowCompDetail.Visible = True
        chkShowVacSickDetail.Visible = True
        
        If chkShowCompDetail Then
            lblCompTimeDt.Visible = True
            dlpDateRange(2).Visible = True
            dlpDateRange(3).Visible = True
        Else
            lblCompTimeDt.Visible = False
            dlpDateRange(2).Visible = False
            dlpDateRange(3).Visible = False
        End If
        
        'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
        If chkShowSickOnly.Visible Then
            chkShowSickOnly = False
            chkShowVacOnly = False
        End If
    Else
        'Hide the Vacation, Sick, Hourly and Comp Time details checkbox
        chkShowHrlyDetail.Visible = False
        chkShowCompDetail.Visible = False
        chkShowVacSickDetail.Visible = False
    
        lblCompTimeDt.Visible = False
        dlpDateRange(2).Visible = False
        dlpDateRange(3).Visible = False
    End If
End Sub

Private Sub chkAtt_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkCompTm_Click(Value As Integer)
    If chkCompTm Then
        chkAtt.Value = False 'Jaddy 11/5/99
        chkCost.Value = False
        If chkSick.Visible Then chkSick = False
        lblCompTimeDt.Visible = True
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
        
        'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
        If chkShowSickOnly.Visible Then
            chkShowSickOnly = False
            chkShowVacOnly = False
        End If
    Else
        lblCompTimeDt.Visible = False
        dlpDateRange(2).Visible = False
        dlpDateRange(3).Visible = False
    End If
End Sub

Private Sub chkCompTm_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkCost_Click(Value As Integer)
    If chkCost Then
        chkAtt.Value = False
        chkCompTm.Value = False
        chkEntitle.Value = False
        If chkSick.Visible Then chkSick = False
        chkShowCalculated.Value = False
        chkShowCalculated.Enabled = False
        
        chkEntBalCost.Visible = True
        dlpDateRange(4).Visible = True
        dlpDateRange(4).Enabled = True
        
        'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
        If chkShowSickOnly.Visible Then
            chkShowSickOnly = False
            chkShowVacOnly = False
        End If
    Else
        chkShowCalculated.Enabled = True
        chkEntBalCost.Visible = False
        dlpDateRange(4).Visible = False
    End If
End Sub

Private Sub chkCost_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkEntBalCost_Click(Value As Integer)
    If chkEntBalCost Then
        dlpDateRange(4).Visible = True
        dlpDateRange(4).Enabled = True
        
        If chkCompTm Then
            dlpDateRange(2).Visible = False
            dlpDateRange(3).Visible = False
        End If
    Else
        dlpDateRange(4).Text = ""
        dlpDateRange(4).Enabled = False
    
        If chkCompTm Then
            dlpDateRange(2).Visible = True
            dlpDateRange(3).Visible = True
        End If
    End If
End Sub

Private Sub chkEntitle_Click(Value As Integer)
    If chkEntitle Then
        chkAtt.Value = False 'Jaddy 11/5/99
        chkCost.Value = False
        If chkSick.Visible Then chkSick = False
        
        'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
        If chkShowSickOnly.Visible Then
            chkShowSickOnly = False
            chkShowVacOnly = False
        End If
    End If
End Sub

Private Sub chkEntitle_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkShowCompDetail_Click(Value As Integer)
    If chkShowCompDetail Then
        lblCompTimeDt.Visible = True
        dlpDateRange(2).Visible = True
        dlpDateRange(3).Visible = True
    Else
        lblCompTimeDt.Visible = False
        dlpDateRange(2).Visible = False
        dlpDateRange(3).Visible = False
    End If
End Sub

Private Sub chkShowSickOnly_Click(Value As Integer)
    'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
    If chkShowSickOnly Then
        chkCompTm = False: chkEntitle = False: chkCost = False 'Jaddy 11/5/99
        chkShowVacOnly = False: chkAtt = False
    End If
End Sub

Private Sub chkShowVacOnly_Click(Value As Integer)
    'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
    If chkShowVacOnly Then
        chkCompTm = False: chkEntitle = False: chkCost = False 'Jaddy 11/5/99
        chkShowSickOnly = False: chkAtt = False
    End If
End Sub

Private Sub chkSick_Click(Value As Integer)
    If chkSick Then
        chkCompTm = False: chkEntitle = False  'Jaddy 11/5/99
        chkAtt = False
    End If
End Sub

Private Sub cmbHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Entitlements Report Criteria", Me) Then Exit Sub
    Screen.MousePointer = HOURGLASS
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdPrint_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
     Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString, , "info:HR"
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdView_GotFocus()
' Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
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
        Case 6: strCd$ = "HREMP.ED_SALDIST"
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

If Len(elpEEID.Text) > 0 Then
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

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpReptau.Text) > 0 Then
    'Ticket #27539 - S.U.C.C.E.S.S. wants to use Reporting Authority 2 instead of 1.
    If glbCompSerial = "S/N - 2422W" Then
        EECri = "{HR_JOB_HISTORY.JH_REPTAU2} IN [" & getEmpnbr(elpReptau.Text) & "] "
    Else
        EECri = "{HR_JOB_HISTORY.JH_REPTAU} IN [" & getEmpnbr(elpReptau.Text) & "] "
    End If
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

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%, X%
Dim FromDate, ToDate, SQLQ
Dim RsHRPARCO As New ADODB.Recordset

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HREMP.ED_EFDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) and "
    
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " ({HREMP.ED_ETDATE} <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Function Cri_SetAll()
Dim X%, xNoFile, xNoWork
Dim WHSCCrpt As Boolean
Dim cmpDates As Boolean
Dim rptTitle As String

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
WHSCCrpt = False

' call cri models set both glbiONeWhere and strSelCri
cmpDates = False

'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_Sup

' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 6
    Call Cri_Code(X%)
Next X%

'Hemu - Muir's Cartage
If glbCompSerial = "S/N - 2340W" Then
    Call Cri_FTDates
End If
'Hemu

If chkAtt Then
    'Ticket #25632 - They have their own custom report so this condition does not work for them
    If glbCompSerial <> "S/N - 2276W" Then  'City of Niagara Falls
        glbstrSelCri = glbstrSelCri & " AND {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    End If
    Call AttDetails
    
    If chkShowCompDetail Then
        Call CompTimeWRK
    End If
End If
    
If chkCompTm Or chkCost Or (glbCompSerial = "S/N - 2276W" And chkAtt) Then  'Ticket #25632
    glbstrSelCri = glbstrSelCri & " AND {HRENTWRK.TE_WRKEMP}='" & glbUserID & "'"
    Call CompTime
    
    If chkEntBalCost Then
        glbstrSelCri = glbstrSelCri & " AND {HRENTCSTWRK.ET_WRKEMP}='" & glbUserID & "'"
        Call EntBalAsOf
        glbstrSelCri = glbstrSelCri & " AND {HRENTHRSWRK.HE_WRKEMP}='" & glbUserID & "'"
        Call HourlyEntitlements_BalAsOf
    End If
End If

'danielk - 12/30/2002 - 7.0 Priority C changes
'Franks - Sep 16,02 for #2823
If glbCompSerial = "S/N - 2297W" Then 'OHTN
    If chkEntitle And cmbHours = "Days" Then
        Call ProjectYTD
    End If
End If
'danielk - 12/30/2002 - end

xNoFile = 7
xNoWork = 99

'report name
If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N - 2236W" Then   'laura 03/09/98 bahamaus
    If comGroup(0).Text <> "(none)" Then
        If chkAtt = True Then          'Jaddy 10/26/99
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentat1.rpt"  'Jaddy 10/26/99
            xNoFile = 5          'Jaddy 10/26/99
        Else          'Jaddy 10/26/99
            If chkEntitle = True And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr5.rpt"
                cmpDates = True
                xNoFile = 8
                xNoWork = 7
            End If
           
            
            If chkEntitle = True And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr2.rpt"
                xNoFile = 7
            End If
            If chkEntitle = False And chkCompTm = True Then
                '----------------------
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentCm1.rpt"
                cmpDates = True
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = False And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthrs.rpt"
                xNoFile = 5
            End If
        End If          'Jaddy 10/26/99
    Else
        If chkAtt = True Then          'Jaddy 10/26/99
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentat2.rpt"  'Jaddy 10/26/99
            xNoFile = 5          'Jaddy 10/26/99
        Else          'Jaddy 10/26/99
            If chkEntitle = True And chkCompTm = True Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr7.rpt"
                cmpDates = True
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = True And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr4.rpt"
                xNoFile = 7
            End If
            If chkEntitle = False And chkCompTm = True Then
                '----------------------
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsentCm3.rpt"
                cmpDates = True
                xNoFile = 8
                xNoWork = 7
            End If
            If chkEntitle = False And chkCompTm = False Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rsenthr1.rpt"
                xNoFile = 5
            End If
        End If      'Jaddy 10/26/99
    End If
  
Else
    'For Rest of the Clients
    
    If chkCost = True Then
        If chkEntBalCost Then
            If comGroup(0).Text <> "(none)" Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr9.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr9a.rpt"
            End If
            If cmbHours.ListIndex = 0 Then
                Me.vbxCrystal.Formulas(9) = "showHoursDays='Hours'"
                rptTitle = "Entitlement Balance (in Hours)/Cost As of " & Format(dlpDateRange(4).Text, "mmm dd, yyyy")
            Else
                Me.vbxCrystal.Formulas(9) = "showHoursDays='Days'"
                rptTitle = "Entitlement Balance (in Days)/Cost As of " & Format(dlpDateRange(4).Text, "mmm dd, yyyy")
            End If
            Me.vbxCrystal.Formulas(8) = "rpttitle='" & rptTitle & "'"
        Else
            If comGroup(0).Text <> "(none)" Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr8.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr8a.rpt"
            End If
            xNoFile = 12
            xNoWork = 7
        End If
    End If
    
    If comGroup(0).Text <> "(none)" Then
        'Grouping NOT NONE
        
        If chkCost = False Then
            If chkShowCalculated.Value = False Then
                Me.vbxCrystal.Formulas(7) = "ShowCalculated='false'"
            Else
                Me.vbxCrystal.Formulas(7) = "ShowCalculated='true'"
            End If
        End If
        
        If cmbHours.ListIndex = 0 Then
            'Hourly Reports
            
            If chkAtt = True Then          'Jaddy 10/26/99
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat1.rpt"  'Jaddy 10/26/99
                xNoFile = 10    '5          'Jaddy 10/26/99
                xNoWork = 5
                
                'Show the Hourly Detail if checkbox checked
                If chkShowHrlyDetail Then
                    Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='true'"
                Else
                    Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='false'"
                End If
                
                'Show the Compensatory Time Detail if checkbox checked
                If chkShowCompDetail Then
                    Me.vbxCrystal.Formulas(6) = "ShowCompDetail='true'"
                Else
                    Me.vbxCrystal.Formulas(6) = "ShowCompDetail='false'"
                End If
            
                'Ticket #25632
                If glbCompSerial = "S/N - 2276W" Then  'City of Niagara Falls
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr5-m2001.rpt"
                    Me.vbxCrystal.Formulas(5) = ""
                    Me.vbxCrystal.Formulas(6) = ""
                    Me.vbxCrystal.Formulas(7) = ""
                End If
            Else          'Jaddy 10/26/99
                If chkEntitle = True And chkCompTm = True Then
                    'Commenting the lines added by Bryan below because this report seems to be
                    'a custom report for unknown client. When we find out the client will uncomment
                    'the code. Otherwise these line are executed by other clients who does not
                    'have this report - hence giving 20507 error
                    'If comGroup(0).Text = "Employee Name" Then 'Added by Bryan 28/M
                    '    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthrC.rpt"
                    '    xNoFile = 11
                    '    xNoWork = 7
                    'Else
                        If glbCompSerial = "S/N - 2276W" Then  'City of Niagara Falls
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr5-m2001.rpt"
                            Me.vbxCrystal.Formulas(7) = ""
                        Else
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr5.rpt"
                        End If
                        cmpDates = True
                        xNoFile = 12  '8 Hemu - report was not working - incomplete assignment
                        xNoWork = 7
                    'End If
                End If
                 
                If chkEntitle = True And chkCompTm = False Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr2.rpt"
                    xNoFile = 7
                End If
                If chkEntitle = False And chkCompTm = True Then
                    '----------------------
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentCm1.rpt"
                    cmpDates = True
                    xNoFile = 12 '9 Hemu - report was not working - incomplete assignment
                    xNoWork = 7
                End If
                If chkEntitle = False And chkCompTm = False And chkCost = False Then
                    If chkSick = True Then 'For North Bay
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentsh1.rpt"
                        xNoFile = 5
                        Me.vbxCrystal.Formulas(7) = ""
                    Else
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthrs.rpt"
                        xNoFile = 5
                    End If
                End If
                
                'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
                If glbCompSerial = "S/N - 2344W" And (chkShowSickOnly = True Or chkShowVacOnly = True) Then
                    If chkShowSickOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344S_rzenthrs.rpt"
                    ElseIf chkShowVacOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344V_rzenthrs.rpt"
                    End If
                End If
            End If          'Jaddy 10/26/99
        Else
            'Days Reports
            
            If chkAtt = True Then          'Jaddy 10/26/99
                If Not glbWHSCC Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat2.rpt"  'Jaddy 10/26/99
                    xNoFile = 10  '5          'Jaddy 10/26/99
                    xNoWork = 5
                    
                    'Show the Hourly Detail if checkbox checked
                    If chkShowHrlyDetail Then
                        Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='true'"
                    Else
                        Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='false'"
                    End If
                    
                    'Show the Compensatory Time Detail if checkbox checked
                    If chkShowCompDetail Then
                        Me.vbxCrystal.Formulas(6) = "ShowCompDetail='true'"
                    Else
                        Me.vbxCrystal.Formulas(6) = "ShowCompDetail='false'"
                    End If
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentat2.rpt"
                    WHSCCrpt = True
                    xNoFile = 5          'Jaddy 10/26/99
                    xNoWork = 5
                    Me.vbxCrystal.Formulas(7) = ""
                End If
            Else          'Jaddy 10/26/99
                If chkEntitle = True And chkCompTm = True Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr6.rpt"
                    cmpDates = True
                    xNoFile = 12
                    xNoWork = 7
                End If
                If chkEntitle = True And chkCompTm = False Then
                    ' danielk - 12/30/2002 - 7.0 Priority C changes
                    ' Franks - Sep 16,02 for #2823
                    If glbCompSerial = "S/N - 2297W" Then 'OHTN
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2297r3.rpt"
                        Me.vbxCrystal.Formulas(7) = ""
                    Else
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr3.rpt"
                    End If
                    ' danielk - 12/30/2002 - end
                    xNoFile = 7
                End If
                If chkEntitle = False And chkCompTm = True Then
                    '----------------------
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentCm2.rpt"
                    cmpDates = True
                    xNoFile = 8
                    xNoWork = 7
                End If
                If chkEntitle = False And chkCompTm = False Then
                    If Not glbWHSCC Then
                        If chkSick = True Then 'For North Bay
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentsd1.rpt"
                            Me.vbxCrystal.Formulas(7) = ""
                        ElseIf chkCost = False Then
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdys.rpt"
                        End If
                    Else
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentdys.rpt"
                        WHSCCrpt = True
                        Me.vbxCrystal.Formulas(7) = ""
                    End If
                    xNoFile = 5
                End If
                
                'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
                If glbCompSerial = "S/N - 2344W" And (chkShowSickOnly = True Or chkShowVacOnly = True) Then
                    If chkShowSickOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344S_rzentdys.rpt"
                    ElseIf chkShowVacOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344V_rzentdys.rpt"
                    End If
                End If
            End If  'Jaddy 10/26/99
        End If
    Else
        'Grouping NONE
        
        If chkCost = False Then
            If chkShowCalculated.Value = False Then
                Me.vbxCrystal.Formulas(7) = "ShowCalculated='false'"
            Else
                Me.vbxCrystal.Formulas(7) = "ShowCalculated='true'"
            End If
        End If
        
        If cmbHours.ListIndex = 0 Then
            'Hourly Reports
            
            If chkAtt = True Then          'Jaddy 10/26/99
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat1.rpt"  'Jaddy 10/26/99
                xNoFile = 10  '5          'Jaddy 10/26/99
                xNoWork = 5
                
                'Show the Hourly Detail if checkbox checked
                If chkShowHrlyDetail Then
                    Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='true'"
                Else
                    Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='false'"
                End If
                
                'Show the Compensatory Time Detail if checkbox checked
                If chkShowCompDetail Then
                    Me.vbxCrystal.Formulas(6) = "ShowCompDetail='true'"
                Else
                    Me.vbxCrystal.Formulas(6) = "ShowCompDetail='false'"
                End If
            Else          'Jaddy 10/26/99
                If chkEntitle = True And chkCompTm = True Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr7.rpt"
                    cmpDates = True
                    xNoFile = 11
                    xNoWork = 7
                End If
                If chkEntitle = True And chkCompTm = False Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr4.rpt"
                    xNoFile = 7
                End If
                If chkEntitle = False And chkCompTm = True Then
                    '----------------------
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentCm3.rpt"
                    cmpDates = True
                    xNoFile = 11
                    xNoWork = 7
                End If
                If chkEntitle = False And chkCompTm = False Then
                    If chkSick = True Then 'For North Bay
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentsh2.rpt"
                        xNoFile = 5
                        Me.vbxCrystal.Formulas(7) = ""
                    ElseIf chkCost = False Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr1.rpt"
                        xNoFile = 8
                    End If
                End If
                
                'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
                If glbCompSerial = "S/N - 2344W" And (chkShowSickOnly = True Or chkShowVacOnly = True) Then
                    If chkShowSickOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344S_rzenthr1.rpt"
                    ElseIf chkShowVacOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344V_rzenthr1.rpt"
                    End If
                End If
            End If          'Jaddy 10/26/99
        Else
            'Days Reports
            
            If chkAtt = True Then          'Jaddy 10/26/99
                If Not glbWHSCC Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentat2.rpt"  'Jaddy 10/26/99
                    xNoFile = 10    '5          'Jaddy 10/26/99
                    xNoWork = 5
                    
                    'Show the Hourly Detail if checkbox checked
                    If chkShowHrlyDetail Then
                        Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='true'"
                    Else
                        Me.vbxCrystal.Formulas(5) = "ShowHrlyDetail='false'"
                    End If
                    
                    'Show the Compensatory Time Detail if checkbox checked
                    If chkShowCompDetail Then
                        Me.vbxCrystal.Formulas(6) = "ShowCompDetail='true'"
                    Else
                        Me.vbxCrystal.Formulas(6) = "ShowCompDetail='false'"
                    End If
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentat2.rpt"
                    WHSCCrpt = True
                    xNoFile = 5          'Jaddy 10/26/99
                    xNoWork = 5
                    
                    Me.vbxCrystal.Formulas(7) = ""
                End If
            Else          'Jaddy 10/26/99
                xNoFile = 5
                If chkEntitle = True And chkCompTm = True Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy7.rpt"
                    cmpDates = True
                    xNoFile = 11
                    xNoWork = 7
                End If
                ' danielk - 12/30/2002 - 7.0 Priority C changes
                If chkEntitle = True And chkCompTm = False Then
                    ' Franks - Sep 16,02 for #2823
                    If glbCompSerial = "S/N - 2297W" Then 'OHTN
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2297Y4.rpt"
                        xNoFile = 7
                        Me.vbxCrystal.Formulas(7) = ""
                    Else
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy4.rpt"
                    End If
                End If
                ' danielk - 12/30/2002 - end
                If chkEntitle = False And chkCompTm = True Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy6.rpt"
                    cmpDates = True
                    xNoFile = 11
                    xNoWork = 7
                End If
                If chkEntitle = False And chkCompTm = False Then
                    If Not glbWHSCC Then
                        If chkSick = True Then 'For North Bay
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentsd2.rpt"
                            Me.vbxCrystal.Formulas(7) = ""
                         ElseIf chkCost = False Then
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzentdy1.rpt"
                        End If
                    Else
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "whentdy1.rpt"
                        WHSCCrpt = True
                        Me.vbxCrystal.Formulas(7) = ""
                    End If
                End If
                
                'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
                If glbCompSerial = "S/N - 2344W" And (chkShowSickOnly = True Or chkShowVacOnly = True) Then
                    If chkShowSickOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344S_rzentdy1.rpt"
                    ElseIf chkShowVacOnly = True Then
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN2344V_rzentdy1.rpt"
                    End If
                End If
            End If 'Jaddy 10/26/99
        End If
    End If
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    Me.vbxCrystal.Formulas(7) = ""
    'If chkAccrued = True Then
    If comGroup(0).Text <> "(none)" Then
        If cmbHours.ListIndex = 0 Then 'Hours
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN23801.rpt"
        Else 'Day
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN23803.rpt"
        End If
    Else
        If cmbHours.ListIndex = 0 Then 'Hours
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN23804.rpt"
        Else 'Day
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "SN23802.rpt"
        End If
    End If
    'End If
End If

If cmpDates = True Then
    vbxCrystal.Formulas(4) = cri_cmpDates
End If
xNoFile = xNoFile + 1

'MsgBox Me.vbxCrystal.ReportFileName

'set to sorting/grouping criteria
X% = Cri_Sorts()   ' returns number of sections formated

If Len(glbstrSelCri) >= 0 Then
    If glbMulti And (glbCompSerial = "S/N - 2428W" Or glbCompSerial = "S/N - 9999W" Or glbCompSerial = "S/N - 2433W") Then
        'Ticket #20490, Ticket #22318 (Kerry's Place)
        glbstrSelCri = glbstrSelCri & " AND ({HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES')"
    End If
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

'set location for database tables
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    If glbWHSCC And WHSCCrpt Then
        If Right(Me.vbxCrystal.ReportFileName, 12) = "whentdy1.rpt" Then
            vbxCrystal.SubreportToChange = "ASLdisp"
        Else
            vbxCrystal.SubreportToChange = "ASL_Disp"
        End If
        vbxCrystal.Connect = RptODBC_SQL
        vbxCrystal.SubreportToChange = "USB_Disp"
        vbxCrystal.Connect = RptODBC_SQL
        vbxCrystal.SubreportToChange = ""
    Else
        'Ticket #25632
        If chkAtt And (Right(Me.vbxCrystal.ReportFileName, 12) = "rzentat1.rpt" Or Right(Me.vbxCrystal.ReportFileName, 12) = "rzentat2.rpt" Or Right(Me.vbxCrystal.ReportFileName, 14) = "RZenthr5-m2001.rpt") Then
            vbxCrystal.SubreportToChange = "RZCompTimeD.rpt"
            vbxCrystal.Connect = RptODBC_SQL
            If Right(Me.vbxCrystal.ReportFileName, 12) = "rzentat1.rpt" Then
                vbxCrystal.SubreportToChange = "RZHrEntH_D.rpt"
            ElseIf Right(Me.vbxCrystal.ReportFileName, 12) = "rzentat2.rpt" Then
                vbxCrystal.SubreportToChange = "RZHrEntD_D.rpt"
            End If
            vbxCrystal.Connect = RptODBC_SQL
            vbxCrystal.SubreportToChange = ""
        End If
    End If
Else
    If glbCompSerial = "S/N - 2297W" And InStr(vbxCrystal.ReportFileName, "2297") <> 0 Then 'OHTN
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To 9
              Me.vbxCrystal.DataFiles(X%) = glbIHRDB
        Next X%
        vbxCrystal.SubreportToChange = "ProjectYTD"
        vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        vbxCrystal.SubreportToChange = ""
    ElseIf xNoWork < 99 Then
        Me.vbxCrystal.Connect = "PWD=petman;"
        For X% = 0 To xNoFile - 1
            If xNoWork = X Then
              Me.vbxCrystal.DataFiles(X%) = glbIHRDBW
            Else
              Me.vbxCrystal.DataFiles(X%) = glbIHRDB
            End If
        Next X%
    Else
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If

End If

'window title if appropriate
Me.vbxCrystal.WindowTitle = "Entitlements Report"

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cri_SetAll", "ENTITLEMENT Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function SAL_SQL()
Dim xStr
Dim strDays As String

'Commented by Bryan 19/Sep/05 - This is flakey code, replacing with something more solid
'If glbSQL Then
'    xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7)), HRPARCO.PC_DECHR) AS SALANN," & _
'    " (CASE WHEN HRPARCO.PC_SERIAL='S/N - 2236W' OR HRPARCO.PC_SERIAL='S/N - 2235W' THEN " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
'    " END) AS SALRATE " & _
'    " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
'ElseIf glbOracle Then
'    xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7)), HRPARCO.PC_DECHR) AS SALANN," & _
'    " (CASE WHEN HRPARCO.PC_SERIAL='S/N - 2236W' OR HRPARCO.PC_SERIAL='S/N - 2235W' THEN " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
'    " END) AS SALRATE " & _
'    " FROM HR_SALARY_HISTORY , HRPARCO WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
'Else
'    xStr = "QRY_SALARYS"
'End If

If GetLeapYear(Year(Date)) Then
    strDays = "366"
Else
    strDays = "355"
End If

xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, "
If glbSQL Or glbOracle Then
   xStr = xStr & "CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * 12), HRPARCO.PC_DECHR) " & _
    " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52 )), HRPARCO.PC_DECHR) " & _
    " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * " & strDays & "), HRPARCO.PC_DECHR) " & _
    " ELSE HR_SALARY_HISTORY.SH_SALARY END AS SALANN,"
    If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N = 2236W" Then
        xStr = xStr & " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
        " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
        " END) AS SALRATE "
    Else
        xStr = xStr & " CASE WHEN HR_SALARY_HISTORY.SH_WHRS =0 THEN 0 ELSE CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * 12) /HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
        " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND(HR_SALARY_HISTORY.SH_SALARY , HRPARCO.PC_DECHR) " & _
        " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52) , HRPARCO.PC_DECHR) " & _
        " ELSE ROUND(HR_SALARY_HISTORY.SH_SALARY* (1/HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
        " END END AS SALRATE "
    End If
Else
    xStr = xStr & "Format(iif(HR_SALARY_HISTORY.SH_SALCD='M',HR_SALARY_HISTORY.SH_SALARY * 12, " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52), " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='D',HR_SALARY_HISTORY.SH_SALARY * " & strDays & ", " & _
    "HR_SALARY_HISTORY.SH_SALARY))),Left('0.0000',(HRPARCO.PC_DECHR+2))) as SALANN, "
    If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N = 2236W" Then
        xStr = "QRY_SALARYS"
    Else
        xStr = xStr & "Format(iif(HR_SALARY_HISTORY.SH_WHRS=0,0,iif(HR_SALARY_HISTORY.SH_SALCD='M',((HR_SALARY_HISTORY.SH_SALARY * 12) /HR_SALARY_HISTORY.SH_WHRS/52), " & _
        "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY, " & _
        "iif(HR_SALARY_HISTORY.SH_SALCD='D',((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52), " & _
        "HR_SALARY_HISTORY.SH_SALARY* (1/HR_SALARY_HISTORY.SH_WHRS/52))))),Left('0.0000',(HRPARCO.PC_DECHR+2))) as SALRATE "
    End If
End If

If glbSQL Then
    xStr = xStr & " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
ElseIf glbOracle Then
    xStr = xStr & " FROM HR_SALARY_HISTORY , HRPARCO WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
ElseIf glbCompSerial <> "S/N - 2235W" And glbCompSerial <> "S/N = 2236W" Then
    xStr = xStr & " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
End If
'added by Bryan on 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    xStr = Replace(xStr, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
End If
SAL_SQL = xStr
End Function

Private Sub CalcCostSQL()
Dim rsEMP As New ADODB.Recordset
Dim SQLQ, SQLX
Dim xtime
On Error GoTo SelCri_Err
Screen.MousePointer = 11
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0
gdbAdoIhr001.Execute "DELETE FROM HRCOEWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"

MDIMain.panHelp(0).FloodPercent = 5


'---------HREMP
SQLQ = " SELECT '001' AS ED_COMPNO,ED_EMPNBR,ED_DIV, ED_DEPTNO,'EDOR' AS ED_ORG_TABL,ED_ORG,ED_PT,"

SQLQ = SQLQ & IIf(IsDate(dlpDateRange(0).Text), Date_SQL(dlpDateRange(0).Text), "Null") & " AS FROMDATE,"
SQLQ = SQLQ & IIf(IsDate(dlpDateRange(1).Text), Date_SQL(dlpDateRange(1).Text), "Null") & " AS TODATE"

SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
'If clpCode(5).Text <> "" Then ' POSITION GROUP
'    SQLQ = SQLQ & " FROM  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB=HRJOB.JB_CODE "
'    SQLQ = SQLQ & " WHERE JB_GRPCD='" & clpCode(5).Text & "'"
'    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND "
'Else
SQLQ = SQLQ & " FROM HREMP WHERE "
'End If
SQLQ = SQLQ & glbSeleDeptUn
If clpDiv.Text <> "" Then SQLQ = SQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then SQLQ = SQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then SQLQ = SQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then SQLQ = SQLQ & "AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then SQLQ = SQLQ & "AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then SQLQ = SQLQ & "AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then SQLQ = SQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then SQLQ = SQLQ & "AND ED_SECTION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(6).Text & "'"
If clpCode(6).Text <> "" Then SQLQ = SQLQ & "AND ED_SALDIST='" & clpCode(6).Text & "'"
If elpEEID.Text <> "" Then SQLQ = SQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then SQLQ = SQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"
If txtShift <> "" Then SQLQ = SQLQ & "AND ED_SHIFT='" & txtShift & "'"
SQLQ = SQLQ & " ORDER BY ED_EMPNBR"

SQLX = "INSERT INTO HRCOEWRK (COMPANY,EMPLOYEE,DIVISION,DEPTNO,UNION_TABL,UNIONS,FT,FROMDATE,TODATE,WRKEMP) "
SQLX = SQLX & in_SQL(glbIHRDBW)
gdbAdoIhr001.Execute SQLX & SQLQ

MDIMain.panHelp(0).FloodPercent = 10
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
xtime = (rsEMP.RecordCount / 500) * 3
If xtime < 2 Then xtime = 2
rsEMP.Close

SQLQ = "SELECT SALS.SH_EMPNBR AS EMPNBR,SUM(SALS.SALANN) AS NUM1, SUM(SALS.SALRATE) AS NUM2, SUM(SALS.SH_WHRS) AS NUM3 "
SQLQ = SQLQ & " FROM " & SAL_SQL() & " AS SALS WHERE SALS.SH_CURRENT<>0 "
If clpCode(5).Text <> "" Then SQLQ = SQLQ & " AND SALS.SH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpCode(5).Text & "')"
SQLQ = SQLQ & " GROUP BY SALS.SH_EMPNBR"

SQLX = " Update HRCOEWRK "
SQLX = SQLX & " SET SALARY = TRSNUM.NUM1, SALRATE=TRSNUM.NUM2, HRSPERWEEK=TRSNUM.NUM3"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 20
'--------HRBENFT
SQLQ = "SELECT BF_EMPNBR AS EMPNBR, SUM(BF_CCOST) AS NUM1 FROM HRBENFT "
SQLQ = SQLQ & " WHERE BF_EDATE<=" & Date_SQL(IIf(IsDate(dlpDateRange(1).Text), dlpDateRange(1).Text, Date))
SQLQ = SQLQ & " GROUP BY BF_EMPNBR"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET BENEFITS = TRSNUM.NUM1"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 30
'--------HREARN
SQLQ = "SELECT EMPNBR,SUM(ACT_DOLLAR) AS NUM1 FROM HREARN WHERE COST_OF_EMPLOYMENT<>0"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " AND TDATE>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND FDATE<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET OTHEREARN = TRSNUM.NUM1"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

   
MDIMain.panHelp(0).FloodPercent = 40
'--------HRTRADE
SQLQ = "SELECT TD_EMPNBR AS EMPNBR,SUM(TD_DUES) AS NUM1 FROM HRTRADE WHERE TD_COMPPD='Y'"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    
    SQLQ = SQLQ & " AND TD_BEGINDT>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND TD_BEGINDT<=" & Date_SQL(dlpDateRange(1).Text)
    
End If
SQLQ = SQLQ & " GROUP BY TD_EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET TRADES = TRSNUM.NUM1"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 50
'--------HREDSEM
SQLQ = "SELECT ES_EMPNBR AS EMPNBR,SUM(ES_TBCO) AS NUM1, SUM(ES_OTHER) AS NUM2, SUM(ES_ACCOM) AS NUM3 FROM HREDSEM "
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " WHERE ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY ES_EMPNBR"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET EDSEM = TRSNUM.NUM1+TRSNUM.NUM2+TRSNUM.NUM3"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 60
'---------HRDOLENT
SQLQ = "SELECT DE_EMPNBR AS EMPNBR,SUM(DE_ENTITLE) AS NUM1 ,SUM(DE_ACTUAL) AS NUM2 FROM HRDOLENT WHERE DE_COST_OF_EMPLOYMENT<>0"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " AND DE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND DE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY DE_EMPNBR"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET DEENTITLE=NUM1,DEDOLLARS = NUM1-NUM2"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 70
'-----------HRENTHRS
SQLQ = "SELECT HE_EMPNBR AS EMPNBR, SUM(HE_ENTITLE) AS NUM1, SUM(HE_TAKEN) AS NUM2 FROM HRENTHRS WHERE HE_COE<>0"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " AND HE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND HE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY HE_EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET HRENTITLE = NUM1*SALRATE,OSHRENTITLE=(NUM1-NUM2)*SALRATE"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 80
'----------HR_ATTENDANCE
SQLQ = "SELECT AD_EMPNBR AS EMPNBR,"

SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS NUM1,"
SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS NUM2,"
SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='WCB' THEN AD_HRS ELSE 0 END) AS NUM3 "

SQLQ = SQLQ & " FROM HR_ATTENDANCE "
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " WHERE AD_DOA>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET OSCOMPTM = (NUM1-NUM2)*SALRATE,WCBDOL=NUM3*SALRATE"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 90
'--------HREMP
SQLQ = "SELECT ED_EMPNBR AS EMPNBR,"
SQLQ = SQLQ & " ED_PVAC AS NUM1,ED_VAC AS NUM2,ED_PSICK AS NUM3,ED_SICK AS NUM4,ED_VACT AS NUM5,ED_SICKT AS NUM6"
SQLQ = SQLQ & " FROM HREMP"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET PVAC = NUM1*SALRATE,VAC=NUM2*SALRATE,PSICK=NUM3*SALRATE,SICK=NUM4*SALRATE,"
SQLX = SQLX & " OSVAC=(NUM1+NUM2-NUM5)*SALRATE,OSSICK = (NUM3+NUM4-NUM6)*SALRATE"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN  (" & SQLQ & ")  AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

 
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SALARY=0 WHERE SALARY IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET HRSPERWEEK=0 WHERE SALRATE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SALRATE=0 WHERE SALRATE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET BENEFITS=0 WHERE BENEFITS IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OTHEREARN=0 WHERE OTHEREARN IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET TRADES=0 WHERE TRADES IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET EDSEM=0 WHERE EDSEM IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSVAC=0 WHERE OSVAC IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSSICK=0 WHERE OSSICK IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET WCBDOL=0 WHERE WCBDOL IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSCOMPTM=0 WHERE OSCOMPTM IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSHRENTITLE=0 WHERE OSHRENTITLE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET DEDOLLARS=0 WHERE DEDOLLARS IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET DEENTITLE=0 WHERE DEENTITLE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET PVAC=0 WHERE PVAC IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET VAC=0 WHERE VAC IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET PSICK=0 WHERE PSICK IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SICK=0 WHERE SICK IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET HRENTITLE=0 WHERE HRENTITLE IS NULL"
MDIMain.panHelp(0).FloodPercent = 100
   
'-------
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
Exit Sub
SelCri_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Costed Employment", "HRCOEWRK", "INSERT")
Resume Next
End Sub

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    If GetEmpData(glbLEE_ID, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(tmpNUM, glbCompDecHR)
End If

End Function

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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
If comGroup(0).Text = lStr("Rept. Authority 1") Then
    grpField$ = "{@fldAdSuper}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
    If comGroup(0).Text = "Employee Name" Then grpField$ = "{@EEFullName}"
End If
grpField$ = Replace(grpField$, "HRTABL.", "HRTABL1.")
If comGroup(0) <> "(none)" Then
    ' <====
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descName1 = " & grpField$
    Me.vbxCrystal.Formulas(1) = dscGroup$
    
    '<====

    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    grpCond$ = "GROUP" & CStr(2) & ";" & "{@EEFullName}" & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$

    Me.vbxCrystal.SectionFormat(0) = "GH1;T;T;X;X;X;X;X"
    
    If chkCost = True And chkEntBalCost = True Then
        'Show the footer of grouping #1 for totals
        Me.vbxCrystal.SectionFormat(1) = "GF1;T,X;X;X;X;X;X"
    Else
        Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"
    End If
Else
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EEFullName}"
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F,F;F;X;X;X;X"
    dscGroup$ = "descGroup1 = ''"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descName1 = ''"
    Me.vbxCrystal.Formulas(1) = dscGroup$
End If
'If chkAtt = False Then  'Jaddy 10/26/99 Jaddy 11/5/99
If chkCost = False Then
    Me.vbxCrystal.Formulas(2) = "DESCSICK =if not isnull({HREMP.ED_EFDATES}) or not isnull({HREMP.ED_ETDATES}) then totext({HREMP.ED_EFDATES}) +' TO ' + totext({HREMP.ED_ETDATES}) "
    Me.vbxCrystal.Formulas(3) = "DESCVAC = if not isnull({HREMP.ED_EFDATE}) or not isnull({HREMP.ED_ETDATE}) then totext({HREMP.ED_EFDATE}) +' TO ' + totext({HREMP.ED_ETDATE}) "
End If
'End If  'Jaddy 10/26/99 Jaddy 11/5/99

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known"), , "info:HR"
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known", , "info:HR"
    'clpDept.SetFocus
    Exit Function
End If


For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpDateRange(2).Text) > 0 Then
    If Not IsDate(dlpDateRange(2).Text) Then
        MsgBox "Compensatory Time 'From' date is invalid"
        dlpDateRange(2).SetFocus
        Exit Function
    End If
End If

If Len(dlpDateRange(3).Text) > 0 Then
    If Not IsDate(dlpDateRange(3).Text) Then
        MsgBox "Compensatory Time 'To' date is invalid"
        dlpDateRange(3).SetFocus
        Exit Function
    End If
End If

If IsDate(dlpDateRange(2).Text) And IsDate(dlpDateRange(3).Text) Then
    If CVDate(dlpDateRange(2).Text) > CVDate(dlpDateRange(3).Text) Then
        MsgBox "Compensatory Time 'From' date cannot be greater than Compensatory Time 'To' date"
        dlpDateRange(2).SetFocus
        Exit Function
    End If
End If

If chkEntBalCost Then
    If Not IsDate(dlpDateRange(4).Text) Then
        MsgBox "When Entitlement Balance/Cost 'As of' is checked then 'As Of Date' is mandatory"
        dlpDateRange(4).SetFocus
        Exit Function
    End If
End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
glbOnTop = "FRMRENTITLE"

Screen.MousePointer = HOURGLASS

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call setRptCaption(Me)
Call setCaption(lblSalDist)
Call setCaption(lblRepAuth)    'Ticket #27539 - It's Reporting Authority 1 from Current Position and not Attendance Supervisor
Call comGrpLoad

'Ticket #27539 - It's Reporting Authority 1 from Current Position and not Attendance Supervisor
'If lblRepAuth.Caption = "AttSupervisor" Then lblRepAuth.Caption = "Supervisor"

'Ticket #27539 - S.U.C.C.E.S.S. wants to use Reporting Authority 2 instead of 1.
If glbCompSerial = "S/N - 2422W" Then
    lblRepAuth.Caption = "Rept. Authority 2"
    Call setCaption(lblRepAuth)
End If

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbLinamar Then clpCode(3).MaxLength = 8

If glbCompSerial = "S/N - 2235W" Then   'laura 03/09/98
    lblTitle(5).Visible = False
    cmbHours.Visible = False
ElseIf glbCompSerial = "S/N - 2236W" Then
    lblTitle(5).Visible = False
    cmbHours.Visible = False
Else
    lblTitle(5).Visible = True
    cmbHours.Visible = True
End If

'For North Bay & District Health Unit Ticket# 7423
If glbCompSerial = "S/N - 2345W" Then
    chkSick.Visible = True
Else
    chkCost.Top = chkSick.Top
    chkEntBalCost.Top = chkSick.Top
    dlpDateRange(4).Top = chkSick.Top
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    'chkAccrued.Visible = True
    chkEntitle.Visible = False
    chkCompTm.Visible = False
    chkAtt.Visible = False
    chkSick.Visible = False
    chkCost.Visible = False
End If

If glbLinamar Then
    lblSalDist.Visible = True
    clpCode(6).Visible = True
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'For Muir's Cartage (Ticket # 5626) - to be made across the board change in the next release
If glbCompSerial = "S/N - 2340W" Then
    lblFromTo.Visible = True
    dlpDateRange(0).Visible = True
    dlpDateRange(1).Visible = True
Else
    lblFromTo.Visible = False
    dlpDateRange(0).Visible = False
    dlpDateRange(1).Visible = False
End If

'Ticket #23878 - Carizon Family and Community Services/KidsLink
If glbCompSerial = "S/N - 2430W" Then
    chkShowCalculated.Caption = "   Show Annual Entitlements"
End If

'Ticket #26481 - Cascade Canada Ltd. - Sick Only or Vacation Only reports
If glbCompSerial = "S/N - 2344W" Then
    chkShowSickOnly.Visible = True
    chkShowVacOnly.Visible = True
Else
    chkShowSickOnly.Visible = False
    chkShowVacOnly.Visible = False
End If

chkEntitle = False

Call INI_Controls(Me)

'Ticket #27095 - Kawartha-Haliburton CAS
If glbCompSerial = "S/N - 2174W" Then
    cmbHours.ListIndex = 0
    chkEntitle.Value = True
End If

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
    If Me.Height >= 8910 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 4000 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 3500
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 9615 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 150
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub


Private Sub CompTime()
Dim SQLQ, SQLX, WSQLQ
On Error GoTo CompTm_Err

SQLQ = "DELETE FROM HRENTWRK " & in_SQL(glbIHRDBW) & " WHERE TE_WRKEMP='" & glbUserID & "'"
SQLQ = SQLQ & " OR TE_WRKEMP = 'NULL' OR (TE_WRKEMP IS NULL) " 'Ticket #14803
gdbAdoIhr001.Execute SQLQ

WSQLQ = WSQLQ & " WHERE " & glbSeleDeptUn
If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If InStr(1, clpCode(1).Text, ",", vbTextCompare) > 1 Then
    WSQLQ = WSQLQ & "AND (ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "'))"
Else
    If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & "AND ED_ORG='" & clpCode(1).Text & "'"
End If

If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & "AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & "AND ED_REGION='" & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & "AND ED_SECTION='" & clpCode(5).Text & "'"
If clpCode(6).Text <> "" Then WSQLQ = WSQLQ & "AND ED_SALDIST='" & clpCode(6).Text & "'"

If elpEEID.Text <> "" Then WSQLQ = WSQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then WSQLQ = WSQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

SQLQ = " SELECT '001' AS TE_COMPNO,AD_EMPNBR,'ADRE' AS TE_REASON_TABL,'CTOT' AS TE_REASON,"
If glbOracle Then
    SQLQ = SQLQ & " SUM(CASE WHEN substr(AD_REASON,1,2)='OT' THEN AD_HRS ELSE 0 END) AS TE_EARNHRS,"
    SQLQ = SQLQ & " SUM(CASE WHEN substr(AD_REASON,1,2)='CT' THEN AD_HRS ELSE 0 END) AS TE_USEDHRS"
ElseIf glbSQL Then
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS TE_EARNHRS,"
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS TE_USEDHRS"
Else
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='OT',AD_HRS,0)) AS TE_EARNHRS, "
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='CT',AD_HRS,0)) AS TE_USEDHRS "
End If

If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        SQLQ = SQLQ & ",'" & glbUserID & "' AS TE_WRKEMP, " & Date_SQL(dlpDateRange(2).Text) & " AS TE_EAFDATE, " & Date_SQL(dlpDateRange(3).Text) & " AS TE_EATDATE"
    ElseIf Len(dlpDateRange(2).Text) > 0 Then
        SQLQ = SQLQ & ",'" & glbUserID & "' AS TE_WRKEMP, " & Date_SQL(dlpDateRange(2).Text) & " AS TE_EAFDATE"
    ElseIf Len(dlpDateRange(3).Text) > 0 Then
        SQLQ = SQLQ & ",'" & glbUserID & "' AS TE_WRKEMP, " & Date_SQL(dlpDateRange(3).Text) & " AS TE_EATDATE"
    End If
Else
    SQLQ = SQLQ & ",'" & glbUserID & "' AS TE_WRKEMP"
    
    If chkEntBalCost Then
        SQLQ = SQLQ & ", " & Date_SQL(dlpDateRange(4).Text) & " AS TE_EATDATE"
    End If
End If
If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE "
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        SQLQ = SQLQ & " AD_DOA >= " & Date_SQL(dlpDateRange(2).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3).Text)
    ElseIf Len(dlpDateRange(2).Text) > 0 Then
        SQLQ = SQLQ & " AD_DOA >= " & Date_SQL(dlpDateRange(2).Text)
    ElseIf Len(dlpDateRange(3).Text) > 0 Then
        SQLQ = SQLQ & " AD_DOA <= " & Date_SQL(dlpDateRange(3).Text)
    End If
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
Else
    SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP " & WSQLQ & ")"
    
    If chkEntBalCost Then
        SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(4).Text)
    End If
End If
SQLQ = SQLQ & " GROUP BY AD_COMPNO,AD_EMPNBR,AD_REASON_TABL "

SQLX = "INSERT INTO HRENTWRK "
If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP,TE_EAFDATE,TE_EATDATE)"
    ElseIf Len(dlpDateRange(2).Text) > 0 Then
        SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP,TE_EAFDATE)"
    ElseIf Len(dlpDateRange(3).Text) > 0 Then
        SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP,TE_EATDATE)"
    End If
Else
    If chkEntBalCost Then
        SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP,TE_EATDATE)"
    Else
        SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP)"
    End If
End If
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & SQLQ
gdbAdoIhr001.Execute SQLX


SQLX = "INSERT INTO HRENTWRK (TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_WRKEMP)"
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & "SELECT ED_COMPNO,ED_EMPNBR,'ADRE' AS AD_REASON_TABL,'SHOW' AS AD_REASON "
SQLX = SQLX & " ,'" & glbUserID & "' AS AD_WRKEMP"
SQLX = SQLX & " FROM HREMP " & WSQLQ
'Add by Franks Dec 4, 2001
'The previues code caused the problem: if both Show Hourly Entitlements and Show Compensatory Time checked,
'Compensatory Time Data not shown, 'SHOW' and Hourly Entitlement code on one can insert in the MDB
'If SHOW' and Hourly Entitlement code both inserted in the MDB, that will cause a problem
SQLX = SQLX & " AND ED_EMPNBR NOT IN (SELECT TE_EMPNBR FROM HRENTWRK "
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & " WHERE HRENTWRK.TE_WRKEMP = '" & glbUserID & "' "
SQLX = SQLX & ") "
'Add by Franks Dec 4, 2001
gdbAdoIhr001.Execute SQLX


Exit Sub

CompTm_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CompTime", "WORK File", "CREATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub

Private Sub AttDetails()

Dim SQLQ, SQLX, WSQLQ

On Error GoTo CompTm_Err

gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"

WSQLQ = glbSeleDeptUn
If clpDiv.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then WSQLQ = WSQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then WSQLQ = WSQLQ & " AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND ED_REGION='" & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then WSQLQ = WSQLQ & " AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SECTION='" & clpCode(5).Text & "'"
If clpCode(6).Text <> "" Then WSQLQ = WSQLQ & " AND ED_SALDIST='" & clpCode(6).Text & "'"
If elpEEID.Text <> "" Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then WSQLQ = WSQLQ & " AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

SQLX = "INSERT INTO HRATTWRK (AD_COMPNO,AD_EMPNBR,AD_REASON_TABL,AD_REASON,AD_WRKEMP)"
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & "SELECT ED_COMPNO,ED_EMPNBR,'ADRE' AS AD_REASON_TABL,'SHOW' AS AD_REASON "
SQLX = SQLX & " ,'" & glbUserID & "' AS AD_WRKEMP"
SQLX = SQLX & " FROM HREMP WHERE " & WSQLQ

gdbAdoIhr001.Execute SQLX

SQLQ = "SELECT AD_COMPNO,AD_EMPNBR,'ADRE' AS AD_REASON_TABL,AD_REASON,AD_DOA,AD_HRS "
SQLQ = SQLQ & " ,'" & glbUserID & "' AS AD_WRKEMP"

If glbOracle Then
    SQLQ = SQLQ & " FROM HR_ATTENDANCE,HREMP WHERE HR_ATTENDANCE.AD_EMPNBR (+) = HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " AND " & WSQLQ
Else
    SQLQ = SQLQ & " FROM HR_ATTENDANCE LEFT JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR = HREMP.ED_EMPNBR"
    SQLQ = SQLQ & " WHERE " & WSQLQ
End If
SQLQ = SQLQ & " AND ((HR_ATTENDANCE.AD_REASON like 'VAC%' "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA>=HREMP.ED_EFDATE "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA<=HREMP.ED_ETDATE) "
SQLQ = SQLQ & " OR  (HR_ATTENDANCE.AD_REASON like 'SIC%' "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA>=HREMP.ED_EFDATES "
SQLQ = SQLQ & " AND HR_ATTENDANCE.AD_DOA<=HREMP.ED_ETDATES))"


SQLX = "INSERT INTO HRATTWRK (AD_COMPNO,AD_EMPNBR,AD_REASON_TABL,AD_REASON,AD_DOA,AD_HRS,AD_WRKEMP) "
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLQ = SQLX & SQLQ


gdbAdoIhr001.Execute SQLQ

Exit Sub

CompTm_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CompTime", "WORK File", "CREATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub

Private Sub comGrpLoad()

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem lStr("Union")
comGroup(0).AddItem lStr("Administered By")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("G/L")
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000

If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If
If Not glbMulti Then comGroup(0).AddItem "Shift"
comGroup(0).AddItem lStr("Rept. Authority 1")
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem "(none)"
comGroup(0).ListIndex = 0
comGroup(1).AddItem "Employee Name"
comGroup(1).ListIndex = 0
comGroup(1).Enabled = False

cmbHours.AddItem "Hours"
cmbHours.AddItem "Days"
cmbHours.ListIndex = 0
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

Private Sub ProjectYTD()
Dim rsVacEnt As New ADODB.Recordset
Dim rsVE As New ADODB.Recordset
Dim snapEntitle As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim SQLQ, SQLX, WSQLQ
Dim VacRule(16, 5), I, K, xOrder, nOrder
Dim xDiv, xDept, xORG, xEMP, xPT, xEDate, xGRPCD
Dim empNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, X%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim xYTD, xYTD0, fglbWDate$, fglbWDate01$

    SQLQ = "DELETE FROM HRENTWRK " & in_SQL(glbIHRDBW)
    gdbAdoIhr001.Execute SQLQ

    Call Pause(3)
    
    SQLQ = "SELECT VE_DIV,VE_DEPT,VE_ORG,VE_EMP,VE_PT,VE_EDATE,VE_GRPCD FROM HRVACENT "
    SQLQ = SQLQ & "GROUP BY VE_DIV,VE_DEPT,VE_ORG,VE_EMP,VE_PT,VE_EDATE,VE_GRPCD"
    If rsVacEnt.State <> 0 Then rsVacEnt.Close
    rsVacEnt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsVacEnt.EOF
        xDiv = "": xDept = "": xORG = "": xEMP = "": xPT = "": xEDate = "": xGRPCD = ""
        If Not IsNull(rsVacEnt("VE_DIV")) Then
            xDiv = rsVacEnt("VE_DIV")
        End If
        If Not IsNull(rsVacEnt("VE_DEPT")) Then
            xDept = rsVacEnt("VE_DEPT")
        End If
        If Not IsNull(rsVacEnt("VE_ORG")) Then
            xORG = rsVacEnt("VE_ORG")
        End If
        If Not IsNull(rsVacEnt("VE_EMP")) Then
            xEMP = rsVacEnt("VE_EMP")
        End If
        If Not IsNull(rsVacEnt("VE_PT")) Then
            xPT = rsVacEnt("VE_PT")
        End If
        If Not IsNull(rsVacEnt("VE_EDATE")) Then
            xEDate = rsVacEnt("VE_EDATE")
        End If
        If Not IsNull(rsVacEnt("VE_GRPCD")) Then
            xGRPCD = rsVacEnt("VE_GRPCD")
        End If
        
        'Get data from HRVACENT
        SQLQ = "SELECT * FROM HRVACENT WHERE (1 =1) "
        If Len(xDiv) > 0 Then
            SQLQ = SQLQ & "AND VE_DIV = '" & xDiv & "' "
        End If
        If Len(xDept) > 0 Then
            SQLQ = SQLQ & "AND VE_DEPT = '" & xDept & "' "
        End If
        If Len(xORG) > 0 Then
            SQLQ = SQLQ & "AND VE_ORG = '" & xORG & "' "
        End If
        If Len(xEMP) > 0 Then
            SQLQ = SQLQ & "AND VE_EMP = '" & xEMP & "' "
        End If
        If Len(xPT) > 0 Then
            SQLQ = SQLQ & "AND VE_PT = '" & xPT & "' "
        End If
        If IsDate(xEDate) Then
        SQLQ = SQLQ & "AND VE_EDATE = " & Date_SQL(xEDate)
        End If
        If Len(xGRPCD) > 0 Then
            SQLQ = SQLQ & "AND VE_GRPCD = '" & xGRPCD & "' "
        End If
        SQLQ = SQLQ & "Order By VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_ORDER "
        
        rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset

        nOrder = 0
        Do While Not rsVE.EOF
            xOrder = rsVE("VE_ORDER")
            nOrder = Format(Val(xOrder), "##0") - 1
            If Not (nOrder < 0 Or nOrder > 16) Then
                I = nOrder
                VacRule(I, 1) = nOrder
                VacRule(I, 2) = rsVE("VE_BMONTH")
                VacRule(I, 3) = rsVE("VE_EMONTH")
                VacRule(I, 4) = rsVE("VE_ENTITLE")
                VacRule(I, 5) = rsVE("VE_TYPE")
            End If
            rsVE.MoveNext
        Loop
        rsVE.Close
        K = nOrder
        
        'Get the employee recordset
        SQLQ = "SELECT HREMP.ED_EMPNBR, qry_JobCurrent.JB_GRPCD, HREMP.ED_VACPC, HREMP.ED_PVAC, HREMP.ED_VAC, HREMP.ED_PSICK, HREMP.ED_SICK, qry_JobCurrent.JH_DHRS, HREMP.ED_DIV, HREMP.ED_EMP, HREMP.ED_DEPTNO, HREMP.ED_PT, HREMP.ED_DOH, HREMP.ED_SENDTE, HREMP.ED_UNION, HREMP.ED_LTHIRE, HREMP.ED_USRDAT1, HREMP.ED_ORG, qry_JobCurrent.JH_FTENUM, qry_JobCurrent.JH_DHRS, HREMP.ED_SECTION, "
        SQLQ = SQLQ & "HREMP.ED_EFDATE, HREMP.ED_ETDATE "
        SQLQ = SQLQ & " FROM HREMP LEFT JOIN qry_JobCurrent ON HREMP.ED_EMPNBR = qry_JobCurrent.JH_EMPNBR "

        SQLQ = SQLQ & "WHERE (1=1) "
        If Len(xDept) > 0 Then SQLQ = SQLQ & " AND  ED_DEPTNO = '" & xDept & "' "
        If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "' "
        If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "' "
        If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "' "
        If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "' "
        If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND JB_GRPCD = '" & xGRPCD & "'; "
        
        If snapEntitle.State <> 0 Then snapEntitle.Close
        snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not snapEntitle.EOF

            empNo& = snapEntitle("ED_EMPNBR")

            If IsNull(snapEntitle("ED_EFDATE")) Then GoTo lblNextRec
            fglbWDate01$ = snapEntitle("ED_EFDATE")
            fglbWDate$ = DateAdd("D", -1, DateAdd("M", 1, fglbWDate01$))
            
            varStartDate = snapEntitle("ED_DOH")
            
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
            
            'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(fglbWDate$)) / 365) * 12
            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(fglbWDate$))
            intWhereFit& = -1   ' first record can be just less than
            xYTD = 0
            For I = 1 To 12
                For X% = 0 To K '16 xYTD
                    If dblServiceYears# >= VacRule(X%, 2) And dblServiceYears# <= VacRule(X%, 3) Then
                        intWhereFit& = X%
                        xYTD0 = VacRule(X%, 4)
                        If VacRule(X%, 5) = "D" Then           ' Entitlements entered in days
                            xYTD = xYTD + xYTD0
                        End If
                        If VacRule(X%, 5) = "H" Then
                            If dblDHours# > 0 Then
                                xYTD = xYTD + (xYTD0 / dblDHours#)
                            End If
                        End If
                        If VacRule(X%, 5) = "F" Then
                            If dblDHours# > 0 Then
                                xYTD = xYTD + (xYTD0 * dblFTEHours#) / dblDHours#
                            End If
                        End If
                    End If
                Next X%
                fglbWDate01$ = DateAdd("M", 1, fglbWDate01$)
                fglbWDate$ = DateAdd("D", -1, DateAdd("M", 1, fglbWDate01$))
                'dblServiceYears# = (DateDiff("d", varStartDate, CVDate(fglbWDate$)) / 365) * 12
                dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(fglbWDate$))
            Next I
            
            If xYTD > 0 Then
                SQLQ = "SELECT * FROM HRENTWRK WHERE TE_EMPNBR = " & empNo&
                If rsWRK.State <> 0 Then rsWRK.Close
                rsWRK.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
                If rsWRK.EOF Then rsWRK.AddNew
                rsWRK("TE_COMPNO") = "001"
                rsWRK("TE_EMPNBR") = empNo&
                rsWRK("TE_EARNHRS") = xYTD
                rsWRK.Update
                rsWRK.Close
            End If
        
lblNextRec:
            
            snapEntitle.MoveNext
        Loop
        
        rsVacEnt.MoveNext
    Loop
    rsVacEnt.Close

End Sub

Private Sub CompTimeWRK()
Dim RsCompTime As New ADODB.Recordset
Dim RsATTC As New ADODB.Recordset
Dim rsPARCO As New ADODB.Recordset
Dim SQLQ, xEmpnbrN, xEmpnbrO, xEarned, xTaken, xBalance, xPaid
Dim xDECHR
Dim WRKSQL As String

WRKSQL = glbSeleDeptUn
If clpDiv.Text <> "" Then WRKSQL = WRKSQL & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then WRKSQL = WRKSQL & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then WRKSQL = WRKSQL & " AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then WRKSQL = WRKSQL & " AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then WRKSQL = WRKSQL & " AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then WRKSQL = WRKSQL & " AND ED_REGION='" & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then WRKSQL = WRKSQL & " AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(5).Text <> "" Then WRKSQL = WRKSQL & " AND ED_SECTION='" & clpCode(5).Text & "'"
If clpCode(6).Text <> "" Then WRKSQL = WRKSQL & " AND ED_SALDIST='" & clpCode(6).Text & "'"
If elpEEID.Text <> "" Then WRKSQL = WRKSQL & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then WRKSQL = WRKSQL & " AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

If Len(dlpDateRange(2).Text) > 0 Or Len(dlpDateRange(3).Text) > 0 Then
    If Len(dlpDateRange(2).Text) > 0 And Len(dlpDateRange(3).Text) > 0 Then
        WRKSQL = WRKSQL & " AND AD_DOA >= " & Date_SQL(dlpDateRange(2).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3).Text)
    ElseIf Len(dlpDateRange(2).Text) > 0 Then
        WRKSQL = WRKSQL & " AND AD_DOA >= " & Date_SQL(dlpDateRange(2).Text)
    ElseIf Len(dlpDateRange(3).Text) > 0 Then
        WRKSQL = WRKSQL & " AND AD_DOA <= " & Date_SQL(dlpDateRange(3).Text)
    End If
End If

xDECHR = 2
rsPARCO.Open "HRPARCO", gdbAdoIhr001, adOpenStatic
If (Not rsPARCO.EOF) And (Not rsPARCO.BOF) Then
    xDECHR = rsPARCO("PC_DECHR")
    If xDECHR < 2 Then xDECHR = 2
End If

SQLQ = "DELETE FROM HRCOMPTIME " & in_SQL(glbIHRDBW) & " WHERE CT_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

If glbOracle Then
    SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
    SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
    SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY,HREMP, HRJOB , HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0 "
    
    If glbMulti Then
        SQLQ = SQLQ & " AND AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
        'SQLQ = SQLQ & " AND AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
        'Ticket #16840 - Only attendance records with current position is retrieved. There could be a job change.
        'SQLQ = SQLQ & " AND AD_JOB = HR_JOB_HISTORY.JH_JOB "
        'SQLQ = SQLQ & " AND AD_JOB = HR_SALARY_HISTORY.SH_JOB "
    End If
    
    If glbCompSerial = "S/N - 2188W" Then
        SQLQ = SQLQ & " AND (SUBSTR(AD_REASON,1,2) = 'OT' OR AD_REASON = 'CT' OR AD_REASON = 'CTP') "   'Hemu
    Else
        SQLQ = SQLQ & " AND (SUBSTR(AD_REASON,1,2) = 'OT' OR SUBSTR(AD_REASON,1,2) = 'CT')"
    End If
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & WRKSQL
    
'    If chkInclAtt Then
'        SQLQ = SQLQ & " UNION "
'        SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
'        SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
'        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY, HR_JOB_HISTORY, HREMP, HRJOB,HR_ATTENDANCE_HISTORY "
'        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
'        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
'        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
'        SQLQ = SQLQ & " AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
'        SQLQ = SQLQ & " AND HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
'        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND SH_CURRENT<>0 "
'
'        If glbMulti Then
'            SQLQ = SQLQ & " AND AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
'            'SQLQ = SQLQ & " AND AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
'            'Ticket #16840 - Only attendance records with current position is retrieved. There could be a job change.
'            'SQLQ = SQLQ & " AND AH_JOB = HR_JOB_HISTORY.JH_JOB "
'            'SQLQ = SQLQ & " AND AH_JOB = HR_SALARY_HISTORY.SH_JOB "
'        End If
'
'        If glbCompSerial = "S/N - 2188W" Then
'            SQLQ = SQLQ & " AND (SUBSTR(AH_REASON,1,2) = 'OT' OR AH_REASON = 'CT' OR AH_REASON = 'CTP')"    'Hemu
'        Else
'            SQLQ = SQLQ & " AND (SUBSTR(AH_REASON,1,2) = 'OT' OR SUBSTR(AH_REASON,1,2) = 'CT')"
'        End If
'        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
'    End If
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

Else

    If glbMulti Then
        SQLQ = "SELECT AD_EMPNBR as SH_EMPNBR,AD_SALARY as SH_SAL,AD_WHRS as SH_WHRS,AD_SALCD as SH_SALCD,AD_SUPER as JH_REPTAU,AD_DHRS as JH_DHRS,"
        SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
        If glbCBrant Then SQLQ = SQLQ & ",AD_CHRGCODE "    'Franks 04/30/03 Ticket# 4079
        SQLQ = SQLQ & " FROM (HR_ATTENDANCE INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
        SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_ATTENDANCE.AD_JOB = HRJOB.JB_CODE)"
        SQLQ = SQLQ & " WHERE 1=1 "
    
        If glbCompSerial = "S/N - 2241W" Then   'Granite Club
            SQLQ = SQLQ & " AND AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
        End If
    Else

        SQLQ = "SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
        SQLQ = SQLQ & " AD_DOA,AD_HRS,AD_REASON,JB_GRPCD "
        If glbCBrant Then SQLQ = SQLQ & ",AD_CHRGCODE "    'Franks 04/30/03 Ticket# 4079
        SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
        SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
        SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
        SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR "
        SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
        
        If glbMulti Then
            SQLQ = SQLQ & " AND AD_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            If glbCompSerial = "S/N - 2241W" Then   'Granite Club
                SQLQ = SQLQ & " AND AD_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
            End If
            'Ticket #16840 - Only attendance records with current position is retrieved. There could be a job change.
            'SQLQ = SQLQ & " AND AD_JOB = HR_JOB_HISTORY.JH_JOB "
            'SQLQ = SQLQ & " AND AD_JOB = HR_SALARY_HISTORY.SH_JOB "
        End If
    End If
    
    SQLQ = SQLQ & " AND (LEFT(AD_REASON,2) = 'OT' OR LEFT(AD_REASON,2) = 'CT')"
    If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & WRKSQL
    
'    If chkInclAtt Then
'        SQLQ = SQLQ & " UNION "
'
'        If glbMulti Then
'            SQLQ = SQLQ & " SELECT AH_EMPNBR as SH_EMPNBR,AH_SALARY as SH_SAL,AH_WHRS as SH_WHRS,AH_SALCD as SH_SALCD,AH_SUPER as JH_REPTAU,AH_DHRS as JH_DHRS,"
'            SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
'            If glbCBrant Then SQLQ = SQLQ & ",AH_CHRGCODE AS AD_CHRGCODE " 'Franks 04/30/03 Ticket# 4079
'            SQLQ = SQLQ & " FROM (HR_ATTENDANCE_HISTORY INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
'            SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_ATTENDANCE_HISTORY.AH_JOB = HRJOB.JB_CODE)"
'            SQLQ = SQLQ & " WHERE 1=1 "
'
'            If glbCompSerial = "S/N - 2241W" Then   'Granite Club
'                SQLQ = SQLQ & " AND AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
'            End If
'
'        Else
'
'            SQLQ = SQLQ & " SELECT SH_EMPNBR,SH_SALARY as SH_SAL,SH_WHRS,SH_SALCD,JH_REPTAU,JH_DHRS,"
'            SQLQ = SQLQ & " AH_DOA AS AD_DOA,AH_HRS AS AD_HRS,AH_REASON AS AD_REASON,JB_GRPCD "
'            If glbCBrant Then SQLQ = SQLQ & ",AH_CHRGCODE AS AD_CHRGCODE " 'Franks 04/30/03 Ticket# 4079
'            SQLQ = SQLQ & " FROM (((HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
'            SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
'            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"
'            SQLQ = SQLQ & " INNER JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR)"
'            SQLQ = SQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE)"
'            SQLQ = SQLQ & " INNER JOIN HR_ATTENDANCE_HISTORY ON HREMP.ED_EMPNBR = HR_ATTENDANCE_HISTORY.AH_EMPNBR "
'            SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND SH_CURRENT<>0 "
'
'            If glbMulti Then
'                SQLQ = SQLQ & " AND AH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
'                If glbCompSerial = "S/N - 2241W" Then   'Granite Club
'                    SQLQ = SQLQ & " AND AH_CHRGCODE = HR_JOB_HISTORY.JH_DEPTNO "
'                End If
'                'Ticket #16840 - Only attendance records with current position is retrieved. There could be a job change.
'                'SQLQ = SQLQ & " AND AH_JOB = HR_JOB_HISTORY.JH_JOB "
'                'SQLQ = SQLQ & " AND AH_JOB = HR_SALARY_HISTORY.SH_JOB "
'            End If
'        End If
'
'        SQLQ = SQLQ & " AND (LEFT(AH_REASON,2) = 'OT' OR LEFT(AH_REASON,2) = 'CT')"
'        If Len(WRKSQL) > 0 Then SQLQ = SQLQ & " AND " & Replace(WRKSQL, "AD_", "AH_")
'    End If
    SQLQ = SQLQ & " ORDER BY SH_EMPNBR,AD_DOA "

End If
'added by Bryan 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    SQLQ = Replace(SQLQ, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
End If

RsATTC.Open SQLQ, gdbAdoIhr001, adOpenStatic

RsCompTime.Open "SELECT * FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' ORDER BY CT_EMPNBR,CT_DOA", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic  ' , adLockBatchOptimistic
Do While Not RsATTC.EOF
    RsCompTime.AddNew
    RsCompTime("CT_EMPNBR") = RsATTC("SH_EMPNBR")
    RsCompTime("CT_REPTAU") = RsATTC("JH_REPTAU")
    RsCompTime("CT_DOA") = RsATTC("AD_DOA")
    RsCompTime("CT_REASON") = RsATTC("AD_REASON")
    RsCompTime("CT_HRS") = RsATTC("AD_HRS")
    RsCompTime("CT_DHRS") = RsATTC("JH_DHRS")
    If RsATTC("SH_WHRS") > 0 Then
        'chanegd by Bryan to include days & months ticket#9354
        If RsATTC("SH_SALCD") = "A" Then
            RsCompTime("CT_HRSSAL") = Round(((RsATTC("SH_SAL") / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
        ElseIf RsATTC("SH_SALCD") = "H" Then
            RsCompTime("CT_HRSSAL") = RsATTC("SH_SAL")
        ElseIf RsATTC("SH_SALCD") = "M" Then
            RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 12) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
        ElseIf RsATTC("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 366) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            Else
                RsCompTime("CT_HRSSAL") = Round((((RsATTC("SH_SAL") * 365) / RsATTC("SH_WHRS")) / 52), xDECHR) '4)
            End If
        End If
    End If
    RsCompTime("CT_GRPCD") = RsATTC("JB_GRPCD")
    RsCompTime("CT_WRKEMP") = glbUserID
    If glbCBrant Then RsCompTime("CT_TEXT1") = RsATTC("AD_CHRGCODE")  'Franks 04/30/03 Ticket# 4079
    RsCompTime.Update
    RsATTC.MoveNext 'CT_HRSSAL
Loop
RsATTC.Close
RsCompTime.Requery

xEmpnbrO = 0

Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("CT_EMPNBR")
    If xEmpnbrO <> xEmpnbrN Then
        xBalance = 0
        xEmpnbrO = xEmpnbrN
    End If
    If Not IsNull(RsCompTime("CT_REASON")) Then
        xEarned = 0
        xTaken = 0
        xPaid = 0
        If Mid(RsCompTime("CT_REASON"), 1, 2) = "OT" Then
            ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
            '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
            '   the OK to remove it.
            'If RsCompTime("CT_HRS") > 0 Then
            xEarned = RsCompTime("CT_HRS")
        End If
        
        If glbCompSerial = "S/N - 2188W" Then   'Hemu
            If RsCompTime("CT_REASON") = "CT" Then
                xTaken = RsCompTime("CT_HRS")
            End If
        Else
            If Mid(RsCompTime("CT_REASON"), 1, 2) = "CT" Then
                ' dkostka - 02/27/2002 - Removed code to make negative amounts zero, not sure
                '   why this was in here, but it shouldn't be.  Checked w/ Linda, she gave
                '   the OK to remove it.
                'If RsCompTime("CT_HRS") > 0 Then
                xTaken = RsCompTime("CT_HRS")
            End If
        End If
        
        If glbCompSerial = "S/N - 2188W" Then   'Hemu
            If RsCompTime("CT_REASON") = "CTP" Then
                xPaid = RsCompTime("CT_HRS")
            End If
        End If
        
        If IsDate(RsCompTime("CT_DOA")) Then
            RsCompTime("CT_WEEK") = WeekdayName(Weekday(RsCompTime("CT_DOA")))
        End If
        RsCompTime("CT_EARN") = xEarned
        RsCompTime("CT_TAKEN") = xTaken
        
        If glbCompSerial = "S/N - 2188W" Then   'Hemu
            RsCompTime("CT_PAID") = xPaid
            xBalance = xBalance + ((xEarned - xTaken) - xPaid)
        Else
            xBalance = xBalance + (xEarned - xTaken)
        End If
        
        RsCompTime("CT_BALANCE") = xBalance
                
        If RsCompTime("CT_DHRS") > 0 Then
            If glbCompSerial = "S/N - 2188W" Then   'Hemu
                RsCompTime("CT_EQUDAYS") = Round((((xEarned - xTaken) - xPaid) / RsCompTime("CT_DHRS")), xDECHR)
            Else
                RsCompTime("CT_EQUDAYS") = Round(((xEarned - xTaken) / RsCompTime("CT_DHRS")), xDECHR)
            End If
        End If
        If RsCompTime("CT_HRSSAL") > 0 Then
            If glbCompSerial = "S/N - 2188W" Then   'Hemu
                RsCompTime("CT_EXTDCOST") = Round((((xEarned - xTaken) - xPaid) * RsCompTime("CT_HRSSAL")), xDECHR)
            Else
                RsCompTime("CT_EXTDCOST") = Round(((xEarned - xTaken) * RsCompTime("CT_HRSSAL")), xDECHR)
            End If
        Else
            RsCompTime("CT_EXTDCOST") = 0
        End If
        RsCompTime.Update
    End If
    RsCompTime.MoveNext
Loop
RsCompTime.Close

If glbCompSerial = "S/N - 2188W" Then   'Hemu
    SQLQ = "SELECT CT_EMPNBR AS EMPNBR, SUM(CT_EARN) AS EARN,SUM(CT_TAKEN) AS TAKEN,SUM(CT_PAID) AS PAID FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' GROUP BY CT_EMPNBR"
Else
    SQLQ = "SELECT CT_EMPNBR AS EMPNBR, SUM(CT_EARN) AS EARN,SUM(CT_TAKEN) AS TAKEN FROM HRCOMPTIME WHERE CT_WRKEMP='" & glbUserID & "' GROUP BY CT_EMPNBR"
End If
RsCompTime.Open SQLQ, gdbAdoIhr001W, adOpenStatic
Do While Not RsCompTime.EOF
    xEmpnbrN = RsCompTime("EMPNBR")
    If glbCompSerial = "S/N - 2188W" Then   'Hemu
        xBalance = (RsCompTime("EARN") - RsCompTime("TAKEN")) - RsCompTime("PAID")
    Else
        xBalance = RsCompTime("EARN") - RsCompTime("TAKEN")
    End If
    SQLQ = "UPDATE HRCOMPTIME " & in_SQL(glbIHRDBW) & " SET CT_BALANCETOT = " & xBalance & " WHERE CT_EMPNBR = " & xEmpnbrN & " AND CT_WRKEMP='" & glbUserID & "'"
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    RsCompTime.MoveNext
Loop

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

Private Sub EntBalAsOf()
    Dim rsHREmp As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsEntCstWrk As New ADODB.Recordset
    Dim SQLQ As String
    Dim xSicTaken As Double
    Dim xVacTaken As Double
    
    'Populate the temp table with Vacation and Sick Time Taken until As Of Date.
    
    'Delete the old records first for this user
    gdbAdoIhr001.Execute "DELETE FROM HRENTCSTWRK " & in_SQL(glbIHRDBW) & " WHERE ET_WRKEMP='" & glbUserID & "'"
    
    SQLQ = "SELECT ED_EMPNBR, ED_PVAC, ED_VAC, ED_PSICK, ED_SICK, ED_EFDATE, ED_EFDATES FROM HREMP"
    SQLQ = SQLQ & " WHERE (ED_EFDATE <= " & Date_SQL(dlpDateRange(4).Text)
    SQLQ = SQLQ & " AND ED_ETDATE >= " & Date_SQL(dlpDateRange(4).Text) & ")"
    SQLQ = SQLQ & " AND (ED_EFDATES <= " & Date_SQL(dlpDateRange(4).Text)
    SQLQ = SQLQ & " AND ED_ETDATES >= " & Date_SQL(dlpDateRange(4).Text) & ")"
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsHREmp.EOF
        'Get total Vacation Taken
        SQLQ = "SELECT SUM(AD_HRS) AS VACTAKEN FROM HR_ATTENDANCE WHERE "
        SQLQ = SQLQ & " AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
        SQLQ = SQLQ & " AND LEFT(AD_REASON,3) = 'VAC' "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(rsHREmp("ED_EFDATE"))
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(4).Text)
        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsAttend.EOF Then
            If Not IsNull(rsAttend("VACTAKEN")) Then
                xVacTaken = rsAttend("VACTAKEN")
            Else
                xVacTaken = 0
            End If
        Else
            xVacTaken = 0
        End If
        rsAttend.Close
        Set rsAttend = Nothing
    
        'Get Total Sick Taken
        SQLQ = "SELECT SUM(AD_HRS) AS SICTAKEN FROM HR_ATTENDANCE WHERE "
        SQLQ = SQLQ & " AD_EMPNBR = " & rsHREmp("ED_EMPNBR")
        SQLQ = SQLQ & " AND LEFT(AD_REASON,3) = 'SIC' "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(rsHREmp("ED_EFDATES"))
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(4).Text)
        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsAttend.EOF Then
            If Not IsNull(rsAttend("SICTAKEN")) Then
                xSicTaken = rsAttend("SICTAKEN")
            Else
                xSicTaken = 0
            End If
        Else
            xSicTaken = 0
        End If
        rsAttend.Close
        Set rsAttend = Nothing
    
        'Add the record in the temp table
        SQLQ = "SELECT * FROM HRENTCSTWRK WHERE 1=2"
        rsEntCstWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsEntCstWrk.AddNew
        rsEntCstWrk("ET_COMPNO") = "001"
        rsEntCstWrk("ET_EMPNBR") = rsHREmp("ED_EMPNBR")
        rsEntCstWrk("ET_ASOF") = dlpDateRange(4).Text
        rsEntCstWrk("ET_VACT") = xVacTaken
        rsEntCstWrk("ET_SICT") = xSicTaken
        rsEntCstWrk("ET_LDATE") = Date
        rsEntCstWrk("ET_LTIME") = Time$
        rsEntCstWrk("ET_LUSER") = glbUserID
        rsEntCstWrk("ET_WRKEMP") = glbUserID
        rsEntCstWrk.Update
        rsEntCstWrk.Close
        Set rsEntCstWrk = Nothing
        
        rsHREmp.MoveNext
    Loop
    rsHREmp.Close
    Set rsHREmp = Nothing
    
End Sub

Private Sub HourlyEntitlements_BalAsOf()
    Dim SQLQ As String
    
    'Delete the old records first for this user
    gdbAdoIhr001.Execute "DELETE FROM HRENTHRSWRK " & in_SQL(glbIHRDBW) & " WHERE HE_WRKEMP='" & glbUserID & "'"
    
    'Populate a temp table with hourly entitlement records and the HE_TAKEN is based on the As Of Date entered and
    'entitlement HE_FDATE
    SQLQ = "INSERT INTO HRENTHRSWRK (HE_COMPNO,HE_EMPNBR,HE_TYPE_TABL,HE_TYPE,HE_FDATE,HE_ENTITLE,HE_COE,HE_DHRS,HE_LDATE"
    SQLQ = SQLQ & ",HE_LTIME,HE_LUSER,HE_PREV,HE_WRKEMP,HE_TDATE,HE_TAKEN)"
    SQLQ = SQLQ & " SELECT HE_COMPNO,HE_EMPNBR,HE_TYPE_TABL,HE_TYPE,HE_FDATE,HE_ENTITLE,HE_COE,HE_DHRS,HE_LDATE,HE_LTIME,HE_LUSER,HE_PREV"
    SQLQ = SQLQ & ",'" & glbUserID & "' AS HE_WRKEMP, " & Date_SQL(dlpDateRange(4).Text) & " AS HE_TDATE"
    SQLQ = SQLQ & ", (SELECT SUM(AD_HRS) FROM HR_ATTENDANCE WHERE AD_EMPNBR = HE_EMPNBR AND AD_REASON = HE_TYPE AND "
    SQLQ = SQLQ & " AD_DOA >= HE_FDATE AND AD_DOA <=" & Date_SQL(dlpDateRange(4).Text) & ") AS HE_TAKEN"
    SQLQ = SQLQ & " FROM HRENTHRS"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRENTHRSWRK SET HE_TAKEN = 0 WHERE HE_TAKEN IS NULL"
    gdbAdoIhr001.Execute SQLQ
End Sub
