VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEEO 
   Caption         =   "EEO Data Maintenance"
   ClientHeight    =   8775
   ClientLeft      =   -555
   ClientTop       =   1485
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   12945
   WindowState     =   2  'Maximized
   Begin VB.TextBox Ethnicity 
      Appearance      =   0  'Flat
      DataField       =   "EO_ETHNICITY"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3075
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbEthnicity 
      Height          =   315
      ItemData        =   "frmEEO.frx":0000
      Left            =   7320
      List            =   "frmEEO.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "Select Ethnicity"
      Top             =   3060
      Width           =   2295
   End
   Begin VB.Frame frmSex 
      Height          =   400
      Left            =   2130
      TabIndex        =   35
      Tag             =   "Gender of Individual"
      Top             =   4440
      Width           =   1815
      Begin Threed.SSOption optSex 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Tag             =   "Gender of Dependent"
         Top             =   120
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Male"
         ForeColor       =   0
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
      Begin Threed.SSOption optSex 
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   9
         Tag             =   "Gender of Dependent"
         Top             =   120
         Width           =   800
         _Version        =   65536
         _ExtentX        =   1411
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Female"
         ForeColor       =   0
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
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEEO.frx":0004
      Left            =   2140
      List            =   "frmEEO.frx":0006
      TabIndex        =   1
      Tag             =   "Type: Applicant or Employee"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox txtEEO 
      Appearance      =   0  'Flat
      DataField       =   "EO_EEONNBR"
      Height          =   285
      Left            =   2130
      TabIndex        =   3
      Tag             =   "11-Enter Emp. Equity No."
      Top             =   2715
      Width           =   1215
   End
   Begin VB.TextBox txtSurname 
      Appearance      =   0  'Flat
      DataField       =   "EO_SURNAME"
      Height          =   285
      Left            =   2130
      TabIndex        =   4
      Tag             =   "Surname"
      Top             =   3075
      Width           =   1215
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      DataField       =   "EO_FNAME"
      Height          =   285
      Left            =   2130
      TabIndex        =   5
      Tag             =   "First Name"
      Top             =   3435
      Width           =   1215
   End
   Begin VB.ComboBox cmbDisability 
      Height          =   315
      ItemData        =   "frmEEO.frx":0008
      Left            =   7315
      List            =   "frmEEO.frx":000A
      TabIndex        =   19
      Tag             =   "Disability--Select Yes or No"
      Top             =   4510
      Width           =   1215
   End
   Begin VB.ComboBox cmbVeteran 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEEO.frx":000C
      Left            =   7315
      List            =   "frmEEO.frx":000E
      TabIndex        =   21
      Tag             =   "Veteran--Select Yes or No"
      Top             =   5270
      Width           =   855
   End
   Begin VB.ComboBox cmbVietnam 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7315
      TabIndex        =   22
      Tag             =   "Vietnam--Select Yes or No"
      Top             =   5630
      Width           =   885
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Tag             =   "EEO Comments"
      Top             =   6480
      Width           =   9495
   End
   Begin VB.TextBox txtGender 
      Appearance      =   0  'Flat
      DataField       =   "EO_SEX"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4498
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EO_DISABLE_YN"
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   10080
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox OETYPE 
      Appearance      =   0  'Flat
      DataField       =   "EO_TYPE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3510
      TabIndex        =   32
      Top             =   1995
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Vietnam 
      Appearance      =   0  'Flat
      DataField       =   "EO_VIETNAM"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5645
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Veteran 
      Appearance      =   0  'Flat
      DataField       =   "EO_VETERAN"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5285
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Disability 
      Appearance      =   0  'Flat
      DataField       =   "EO_DISABLE_YN"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4515
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
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
      Left            =   4320
      TabIndex        =   25
      Tag             =   "Find Employee"
      Top             =   7530
      Width           =   735
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "Sort by Employee Number"
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
      Left            =   5400
      TabIndex        =   26
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   7530
      Width           =   2415
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   24
      Tag             =   "00-Search for Surname"
      Top             =   7575
      Width           =   1935
   End
   Begin VB.ComboBox comCountryOfEmp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2130
      TabIndex        =   11
      Tag             =   "00-Country of Employment"
      Top             =   5270
      Width           =   1320
   End
   Begin VB.TextBox txtCountryOfEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "EO_WORKCOUNTRY"
      Height          =   285
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   28
      TabStop         =   0   'False
      Tag             =   "01-Country"
      Top             =   5285
      Visible         =   0   'False
      Width           =   555
   End
   Begin INFOHR_Controls.CodeLookup clpNOGC 
      DataField       =   "EO_OCC_CAT"
      Height          =   285
      Left            =   7005
      TabIndex        =   18
      Tag             =   "Enter Job Category Code"
      Top             =   4155
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   6
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EO_DISABLE"
      Height          =   285
      Index           =   2
      Left            =   7005
      TabIndex        =   20
      Tag             =   "Select Disability Code"
      Top             =   4900
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDDI"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EO_RACE"
      Height          =   285
      Index           =   1
      Left            =   7005
      TabIndex        =   16
      Tag             =   "Enter a EEO Race Code"
      Top             =   3435
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRC"
   End
   Begin INFOHR_Controls.DateLookup dlpDOB 
      DataField       =   "EO_DOB"
      Height          =   285
      Left            =   1815
      TabIndex        =   7
      Tag             =   "41-Enter Date of Birth"
      Top             =   4155
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmEEO.frx":0010
      Height          =   1635
      Left            =   120
      OleObjectBlob   =   "frmEEO.frx":0024
      TabIndex        =   0
      Tag             =   "Department Listings"
      Top             =   120
      Width           =   10680
   End
   Begin MSMask.MaskEdBox txtSIN 
      DataField       =   "EO_SSN"
      Height          =   285
      Left            =   2130
      TabIndex        =   6
      Tag             =   "Social Security Number"
      Top             =   3795
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###-##-####"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      DataField       =   "EO_EMPNBR"
      Height          =   285
      Left            =   1815
      TabIndex        =   2
      Tag             =   "10-Enter Employee Number"
      Top             =   2355
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   3040
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDOH 
      DataField       =   "EO_DOH"
      DataSource      =   " "
      Height          =   285
      Left            =   1815
      TabIndex        =   10
      Tag             =   "41-Original Hire Date "
      Top             =   4900
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1060
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      DataField       =   "EO_PT"
      Height          =   285
      Left            =   7005
      TabIndex        =   12
      Tag             =   "EDPT-Category"
      Top             =   1995
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
      MaxLength       =   7
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EO_REGION"
      Height          =   285
      Index           =   0
      Left            =   7005
      TabIndex        =   13
      Tag             =   "00-Region"
      Top             =   2355
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EO_LOC"
      Height          =   285
      Index           =   3
      Left            =   7005
      TabIndex        =   14
      Tag             =   "00-Location - Code"
      Top             =   2715
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   59
      Top             =   8355
      Width           =   12945
      _Version        =   65536
      _ExtentX        =   22834
      _ExtentY        =   741
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
      Begin VB.CommandButton cmdActTerm 
         Caption         =   "Terminated EEO"
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
         Left            =   2160
         TabIndex        =   63
         Top             =   0
         Width           =   1875
      End
      Begin VB.CommandButton cmdRecal 
         Caption         =   "&Recalculate"
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
         Left            =   360
         TabIndex        =   27
         Top             =   0
         Width           =   1275
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   6600
         Top             =   120
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8595
         Top             =   45
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EO_RACE2"
      Height          =   285
      Index           =   4
      Left            =   7005
      TabIndex        =   17
      Tag             =   "Enter a EEO Race Code"
      Top             =   3795
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRC"
   End
   Begin MSMask.MaskEdBox medsalary 
      Height          =   285
      Left            =   7320
      TabIndex        =   67
      Tag             =   "21-Enter salary"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   -2147483633
      Enabled         =   0   'False
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTermDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "DOT"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd-MMMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   66
      Top             =   6000
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Salary"
      Height          =   195
      Index           =   19
      Left            =   5280
      TabIndex        =   65
      Top             =   6000
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Termination Date"
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   64
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Ethnicity"
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
      Index           =   17
      Left            =   5280
      TabIndex        =   62
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Race 2"
      Height          =   195
      Index           =   16
      Left            =   5280
      TabIndex        =   60
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   58
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   57
      Top             =   2400
      Width           =   1290
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "EEO Number"
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
      Index           =   2
      Left            =   120
      TabIndex        =   56
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
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
      Index           =   3
      Left            =   120
      TabIndex        =   55
      Top             =   3120
      Width           =   750
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
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
      Index           =   4
      Left            =   120
      TabIndex        =   54
      Top             =   3480
      Width           =   915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Social Security Number"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   53
      Top             =   3840
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      Index           =   6
      Left            =   120
      TabIndex        =   52
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Index           =   7
      Left            =   120
      TabIndex        =   51
      Top             =   4570
      Width           =   630
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Race"
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
      Index           =   8
      Left            =   5280
      TabIndex        =   50
      Top             =   3480
      Width           =   465
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "NOC Code"
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
      Index           =   9
      Left            =   5280
      TabIndex        =   49
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Disability"
      Height          =   195
      Index           =   10
      Left            =   5280
      TabIndex        =   48
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Disability Code"
      Height          =   195
      Index           =   11
      Left            =   5280
      TabIndex        =   47
      Top             =   4945
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Veteran"
      Height          =   195
      Index           =   12
      Left            =   5280
      TabIndex        =   46
      Top             =   5330
      Width           =   555
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Vietnam Era"
      Height          =   195
      Index           =   13
      Left            =   5280
      TabIndex        =   45
      Top             =   5690
      Width           =   915
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   44
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label lblCNum 
      AutoSize        =   -1  'True
      Caption         =   "Comp"
      DataField       =   "EO_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3810
      TabIndex        =   43
      Top             =   3600
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EO_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3810
      TabIndex        =   42
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Surname"
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
      TabIndex        =   41
      Top             =   7620
      Width           =   1665
   End
   Begin VB.Label lblOHire 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Original Hire"
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
      TabIndex        =   40
      Top             =   4945
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Employment"
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
      Index           =   15
      Left            =   120
      TabIndex        =   39
      Top             =   5330
      Width           =   1950
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5280
      TabIndex        =   38
      Top             =   2040
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   5280
      TabIndex        =   37
      Top             =   2400
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   5280
      TabIndex        =   36
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "frmEEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RFound As Integer
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew
Dim EEO_snap As New ADODB.Recordset 'EEO_Snap As snapshot
Dim NOC_Snap As New ADODB.Recordset
Dim Race_Snap As New ADODB.Recordset
Dim Disa_Snap As New ADODB.Recordset
Dim HR_Snap As New ADODB.Recordset
Dim xActn
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim EESNameSort As Integer
Dim EEPIDSort As Boolean
Dim xActTermFlag As String

Sub cmbDisability_Click()
clpCode(2).Enabled = (cmbDisability.Text = "Yes")
If cmbDisability = "No" Then
    clpCode(2).Text = ""
End If
End Sub
Sub cmbDisability_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbEthnicity_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbType_Change()
Call cmbType_Click
End Sub

Sub cmbType_Click()
If cmbType.ListIndex = 0 Then
    elpEEID.ShowDescription = True
    'lblEEName.Visible = True
    OETYPE = "E"
    lblOHire.Caption = "Original Hire"
Else
    OETYPE = "A"
    lblOHire.Caption = "Date of Application"
End If
'elpEEID.text = ""
txtSurname.Enabled = OETYPE = "A"
txtFName.Enabled = OETYPE = "A"
txtSIN.Enabled = OETYPE = "A"
dlpDOB.Enabled = OETYPE = "A"
frmSex.Enabled = OETYPE = "A"
elpEEID.Enabled = Not OETYPE = "A"
End Sub

Sub cmbType_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbVeteran_Click()
    cmbVietnam.Enabled = (cmbVeteran.Text = "Yes")
End Sub

Sub cmbVeteran_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmbVietnam_Click()

If cmbVietnam.ListIndex = 0 Then
    Vietnam = True
Else
    Vietnam = False
End If

End Sub

Sub cmbVietnam_GotFocus()
    Call SetPanHelp(ActiveControl)
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
'
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
    'cmdModify.Enabled = False
    'cmdDelete.Enabled = False
    'cmdPrint.Enabled = False
End If

cmbType.Enabled = TF
elpEEID.Enabled = TF
txtEEO.Enabled = TF
txtSurname.Enabled = TF
txtFName.Enabled = TF
txtSIN.Enabled = TF
dlpDOB.Enabled = TF
frmSex.Enabled = TF
 clpCode(1).Enabled = TF
 clpCode(2).Enabled = TF
 clpNOGC.Enabled = TF
cmbDisability.Enabled = TF
cmbVeteran.Enabled = TF
cmbVietnam.Enabled = TF
txtComments.Enabled = TF
'vbxTrueGrid.Enabled = FT

'Ticket #22682 - Release 8.0
cmbEthnicity.Enabled = TF
clpCode(4).Enabled = TF

'Ticket #23947 Franks 06/20/2013 - begin
If Not gSec_Show_SIN_SSN Then
    txtSIN.Visible = False
End If
If Not gSec_Show_DOB Then
    dlpDOB.Visible = False
End If
'Ticket #23947 Franks 06/20/2013 - end
End Sub

Public Sub cmdCancel_Click()

On Error GoTo Can_Err


rsDATA.CancelUpdate
Call Display_Value


xActn = ""

Call cmdModify_Click
Call SET_UP_MODE  ' reset screen's attributes

elpEEID.Enabled = False
txtEEO.Enabled = False

fglbNew = False

xActn = ""
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRMATRIX", "Cancel")
Call RollBack

End Sub
Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdClose_Click()
Call NextForm 'Ticket #30482 Franks 08/16/2017
Unload Me

End Sub

Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

Call SET_UP_MODE
elpEEID.Enabled = False
txtEEO.Enabled = False


Me.vbxTrueGrid.SetFocus

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRMATRIX", "Delete")
Call RollBack

End Sub

Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdModify_Click()

On Error GoTo Mod_Err
'Data1.Recordset.Edit
Call ST_UPD_MODE(True)
cmbType.Enabled = False
elpEEID.Enabled = False
txtEEO.Enabled = False

Call cmbDisability_Click
Call cmbVeteran_Click

'vbxTrueGrid.Enabled = False
xActn = "C"

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRMATRIX", "Modify")
Call RollBack

End Sub

Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdNew_Click()

On Error GoTo AddN_Err

fglbNew = True

'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)
rsDATA.AddNew

xActn = "A"

Call SET_UP_MODE

cmbType.ListIndex = 0
cmbDisability.ListIndex = 1
optSex(0) = True 'False
optSex(1) = False

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

If Err = 3021 Then
    'Data1.RecordSource = "HREEO"
    fglbEmptyNew = True
    Data1.Refresh
    Data1.Recordset.AddNew
    Resume Next
End If

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRMATRIX", "Add")
Call RollBack

End Sub

Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdOK_Click()
Dim rs As New ADODB.Recordset
Dim x%
Dim xChange1, xChange2
On Error GoTo cmdOK_Err

Vietnam = IIf(cmbVietnam.Text = "Yes", "-1", "0")
Veteran = IIf(cmbVeteran.Text = "Yes", "-1", "0")
Disability = IIf(cmbDisability.Text = "Yes", "-1", "0")

If Not chkEEO() Then Exit Sub

'Ticket #22682 - Release 8.0 - Ethnicity
'Ethnicity = IIf(cmbEthnicity.Text = "Hispanic or Latino", "HL", "NHL")
If cmbEthnicity.Text = "Hispanic or Latino" Then
    Ethnicity = "HL"
ElseIf cmbEthnicity.Text = "Not Hispanic or Latino" Then
    Ethnicity = "NHL"
Else
Ethnicity = ""
End If

lblCNum = "001"
rsDATA!EO_SSN = Format(txtSIN, "#########")

Call Set_Control("U", Me, rsDATA)

rsDATA!EO_DISABLE_YN = 0

If cmbType.Text = "Employee" Then
    rsDATA!EO_EMPNBR = elpEEID 'Ticket# 7867
End If

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh

Call cmdModify_Click

fglbNew = False

Call SET_UP_MODE
vbxTrueGrid.SetFocus

elpEEID.Enabled = False
txtEEO.Enabled = False

Screen.MousePointer = DEFAULT

Call NextForm 'Ticket #30482 Franks 08/16/2017

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREEO", "Update")
Call RollBack

End Sub

Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String
Dim xReport
RHeading = "EEO DATA"

Me.vbxCrystal.Reset 'Ticket #23947 Franks 06/20/2013
'Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If xActTermFlag = "ACT" Then
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    xReport = glbIHRREPORTS & "rgrideeo.rpt"
Else
    Me.vbxCrystal.WindowTitle = "" & RHeading & " Report for Terminated Employees"
    xReport = glbIHRREPORTS & "rgrideeoT.rpt" 'Ticket #25836 Franks 08/05/2014
End If
Me.vbxCrystal.ReportFileName = xReport

'Ticket #23947 Franks 06/20/2013 - begin
glbstrSelCri = ""
glbiOneWhere = False
Call glbCri_DeptUN("")
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
Me.vbxCrystal.Formulas(10) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
Me.vbxCrystal.Formulas(11) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
'Ticket #23947 Franks 06/20/2013 - end

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    'Me.vbxCrystal.DataFiles(2) = glbIHRDB
    'Me.vbxCrystal.DataFiles(3) = glbIHRDB
End If

Me.vbxCrystal.Action = 1
End Sub

Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdView_Click()
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    cmdPrint_Click
End Sub

Private Sub cmdActTerm_Click()
    If xActTermFlag = "ACT" Then
        cmdActTerm.Caption = "Active EEO"
        xActTermFlag = "TERM"
        Call ShowTermDateRate(True)
        Call EERetrieve
        Call Display_Value
        Exit Sub
    End If
    If xActTermFlag = "TERM" Then
        cmdActTerm.Caption = "Terminated EEO"
        xActTermFlag = "ACT"
        Call ShowTermDateRate(False)
        Call EERetrieve
        Call Display_Value
        Exit Sub
    End If
    
End Sub

Private Sub ShowTermDateRate(xFlag)
lblTitle(18).Visible = xFlag
lblTitle(19).Visible = xFlag
lblTermDate.Visible = xFlag
medsalary.Visible = xFlag
medsalary.BorderStyle = 0 ' mskNone
'medsalary.Enabled = True 'False
If xFlag Then
    lblTermDate.DataField = "TERM_DOT"
    medsalary.DataField = "SH_SALARY"
    vbxTrueGrid.Columns(11).Visible = True
    vbxTrueGrid.Columns(12).Visible = True
Else
    lblTermDate.DataField = ""
    medsalary.DataField = ""
    vbxTrueGrid.Columns(11).Visible = False
    vbxTrueGrid.Columns(12).Visible = False
End If
End Sub

Private Sub cmdEESort_Click()
Dim xStr
txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

EEPIDSort = False

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort.Caption = "Sort by Surname "
    glbSort = "NUMBER"
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort.Caption = "Sort by Employee Number"
    glbSort = "NAME"
End If

If EEList() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = lblSearchBy.Caption '"Search by Surname "    'laura jan 05 1998

End Sub

Private Function EEList()
Dim SQLQ As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

On Error GoTo EEList_Err
EEList = False
SQLQ = "Select * FROM HREEO "
SQLQ = SQLQ & "WHERE EO_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ") "

If cmdEESort.Caption = "Sort by Employee Number" Then
    SQLQ = SQLQ & " ORDER BY EO_SURNAME, EO_FNAME"
Else
    SQLQ = SQLQ & " ORDER BY EO_EMPNBR"
End If

Data1.RecordSource = SQLQ
Data1.Refresh
Me.vbxTrueGrid.Refresh

EEList = True

Exit Function

EEList_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EEList", "HREMP", "Select")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark
On Error GoTo Srch_Err
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch, "'", "''")

    If EESNameSort = True Then
        SQLQ = "EO_SURNAME  >= '" & Sch & "'"
    Else
        If Not IsNumeric(txtEESearch) Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        SQLQ = "EO_EMPNBR >= '" & Sch & "'"
    End If

    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    If Data1.Recordset.RecordCount > 0 Then Data1.Recordset.MoveFirst
    MsgBox "Employee not found"
End If
Screen.MousePointer = DEFAULT

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EEList", "HREEO", "Find Next")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdRecal_Click()
Dim Msg As String, a%

    'Msg = "This program will updates NOC and Termination Date if missing or incorrect."
    'Msg = Msg & Chr(10) & "Are you sure you want to do it?"
    Msg = "Are you sure you want to re-update the NOC Code and the employee demographics data?"
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 1
    MDIMain.panHelp(0).FloodPercent = 3

    'Call UpdateHREEO_NOGC
    'Call InputHREEO_DOT
    Call uptEEO_Fields("", "Update")
    
    Data1.Refresh
    
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
End Sub

'Frank 07/22/2010  Ticket #18790, Update Date of NOGC
Private Sub UpdateHREEO_NOGC()
On Error GoTo UpdateHREEO_NOGC_Err
Dim rsEmpEEO As New ADODB.Recordset
Dim rsEmpNOC As New ADODB.Recordset
Dim SQLQ As String
Dim dblPerc, FloodPerc As Double
    
    SQLQ = "SELECT * FROM HREEO WHERE NOT (EO_EMPNBR IS NULL)"
    rsEmpEEO.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsEmpEEO.EOF Then
        rsEmpEEO.MoveFirst
        
        dblPerc = (50 / rsEmpEEO.RecordCount)
        FloodPerc = dblPerc
        
        gdbAdoIhr001.BeginTrans
        Do While Not rsEmpEEO.EOF
            rsEmpNOC.Open "SELECT JB_FEDGRP FROM HRJOB WHERE JB_CODE = (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & rsEmpEEO("EO_EMPNBR") & " AND JH_CURRENT <> 0)", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            If Not rsEmpNOC.EOF Then
                If Not IsNull(rsEmpNOC("JB_FEDGRP")) Then
                    gdbAdoIhr001.Execute "UPDATE HREEO SET EO_OCC_CAT = '" & rsEmpNOC("JB_FEDGRP") & "' WHERE EO_EMPNBR = " & rsEmpEEO("EO_EMPNBR")
                End If
            End If
            rsEmpEEO.MoveNext
            
            MDIMain.panHelp(0).FloodPercent = FloodPerc
            FloodPerc = FloodPerc + dblPerc
            
            rsEmpNOC.Close
        Loop
        gdbAdoIhr001.CommitTrans
        
    End If
    rsEmpEEO.Close
    MDIMain.panHelp(0).FloodPercent = 50
    Exit Sub
    
    
UpdateHREEO_NOGC_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateHREEO_NOGC", "HREEO", "Update")
Call RollBack  '08June99 js

End Sub

'Frank 07/22/2010  Ticket #18790, Update Date of Termination
Private Sub InputHREEO_DOT()
On Error GoTo InputHREEO_DOT_Err

Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset
Dim rsTermEmp As New ADODB.Recordset
Dim dblPerc, FloodPerc As Double

SQLQ = "SELECT * FROM HREEO WHERE NOT (EO_EMPNBR IS NULL)"
dynEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Not dynEmp.EOF Then
    dynEmp.MoveFirst
    
    dblPerc = (50 / dynEmp.RecordCount)
    FloodPerc = 50 + dblPerc
    
    gdbAdoIhr001.BeginTrans
    Do While Not dynEmp.EOF
        rsTermEmp.Open "SELECT Term_DOT FROM TERM_HRTRMEMP WHERE Employee_Number = " & dynEmp("EO_EMPNBR"), gdbAdoIhr001X, adOpenStatic
        
        If Not rsTermEmp.EOF Then
            gdbAdoIhr001.Execute "UPDATE HREEO SET EO_DOT = " & Date_SQL(rsTermEmp("Term_DOT")) & " WHERE EO_EMPNBR = " & dynEmp("EO_EMPNBR")
        End If
        dynEmp.MoveNext
        
        MDIMain.panHelp(0).FloodPercent = FloodPerc
        FloodPerc = FloodPerc + dblPerc
        
        rsTermEmp.Close
    Loop
    gdbAdoIhr001.CommitTrans
    
End If
dynEmp.Close
Exit Sub

InputHREEO_DOT_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InputHREEO_DOT", "TERM_HRTRMEMP", "Update")
Call RollBack  '08June99 js

End Sub


Private Sub comCountryOfEmp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comCountryOfEmp_LostFocus()
txtCountryOfEmp.Text = comCountryOfEmp.Text
End Sub

Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "data1.error", "DataSource", "")

End Sub

Private Sub EERetrieve()
Dim SQLQ

If xActTermFlag = "ACT" Then 'Ticket #25669 Franks 06/24/2014
    'Frank 06/17/2004 Ticket# 6380, no security control before
    SQLQ = "SELECT * FROM HREEO "
    SQLQ = SQLQ & "WHERE (EO_TYPE = 'E' and EO_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")) OR EO_TYPE = 'A' "
    SQLQ = SQLQ & "ORDER BY EO_SURNAME "
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = SQLQ
    Data1.Refresh
End If

If xActTermFlag = "TERM" Then 'Ticket #25669 Franks 06/24/2014
    SQLQ = "Select Term_HREEO.*, TERM_DOT,SH_SALARY "
    SQLQ = SQLQ & " FROM Term_HREEO "
    SQLQ = SQLQ & "LEFT JOIN Term_HRTRMEMP ON Term_HREEO.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ "
    SQLQ = SQLQ & "LEFT JOIN Term_SALARY_HISTORY ON Term_HREEO.TERM_SEQ=Term_SALARY_HISTORY.TERM_SEQ "
    SQLQ = SQLQ & "WHERE (EO_TYPE = 'E' and Term_HREEO.TERM_SEQ IN (SELECT TERM_SEQ FROM Term_HREMP WHERE " & glbSeleDeptUn & ")) OR EO_TYPE = 'A' "
    SQLQ = SQLQ & " AND NOT Term_SALARY_HISTORY.SH_CURRENT = 0 "
    SQLQ = SQLQ & "ORDER BY EO_SURNAME "
    
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = SQLQ
    Data1.Refresh
End If

End Sub

Function EEIDRetrieve()

Dim SQLQ As String

EEIDRetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

SQLQ = "Select HREMP.* "
SQLQ = SQLQ & " FROM HREMP"
SQLQ = SQLQ & " WHERE HREMP.ED_EMPNBR = " & glbLEE_ID & ";"

EEIDRetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DEPRetrieve", "HRDEPEND", "SELECT")
Call RollBack

Exit Function

End Function

Private Sub Disability_Change()
    cmbDisability.Text = IIf(Disability.Text = "0", "No", "Yes")
End Sub

Private Function CR_FLNames()
Dim SQLQ As String
Dim countr   As Integer
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset

On Error GoTo CR_FLNames_Err

SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & elpEEID.Text & " "
rsTB.Open SQLQ, gdbAdoIhr001, adOpenStatic
'rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
'rsTB.MoveFirst
'rsTB.Find "ED_EMPNBR = " & Val(elpEEID)
If Not rsTB.EOF Then
    txtSurname = rsTB("ED_Surname")
    txtFName = rsTB("ED_FName")
    txtSIN = rsTB("ED_SIN")
    dlpDOB = rsTB("ED_DOB")
    txtGender = rsTB("ED_SEX")
    If txtGender = "M" Then
        optSex(0) = True
    Else
        optSex(1) = True
    End If
    'Ticket #18790 - begin
    If Not IsNull(rsTB("ED_DOH")) Then dlpDOH.Text = rsTB("ED_DOH")
    If Not IsNull(rsTB("ED_WORKCOUNTRY")) Then txtCountryOfEmp.Text = rsTB("ED_WORKCOUNTRY")
    If Not IsNull(rsTB("ED_PT")) Then clpPT.Text = rsTB("ED_PT")
    If Not IsNull(rsTB("ED_REGION")) Then clpCode(0).Text = rsTB("ED_REGION")
    If Not IsNull(rsTB("ED_LOC")) Then clpCode(3).Text = rsTB("ED_LOC")
    'NOC code
    clpNOGC.Text = getHREEO_NOC(elpEEID.Text)
    'Ticket #18790 - end
End If
rsTB.Close
Exit Function

CR_FLNames_Err:
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_FLNames", "HREMP", "Select")
Call RollBack

End Function

Private Sub elpEEID_LostFocus()
    If Len(elpEEID) > 0 And Len(txtEEO) = 0 Then
        txtEEO = elpEEID
        Call CR_FLNames
    End If
End Sub

Private Sub Ethnicity_Change()
    'cmbEthnicity.Text = IIf(Ethnicity.Text = "HL", "Hispanic or Latino", "Not Hispanic or Latino")
    If Ethnicity.Text = "HL" Then
        cmbEthnicity.Text = "Hispanic or Latino"
    ElseIf Ethnicity.Text = "NHL" Then
        cmbEthnicity.Text = "Not Hispanic or Latino"
    Else
        cmbEthnicity.ListIndex = -1
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus
'Call cmdModify_Click
glbOnTop = "FRMEEO"
elpEEID.Enabled = False
txtEEO.Enabled = False

End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEEO"
End Sub

Private Sub Form_Load()
Dim SQLQ
glbOnTop = "FRMEEO"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

xActTermFlag = "ACT" 'Ticket #25669 Franks 06/24/2014

Call ShowTermDateRate(False)

Call EERetrieve

fglbNew = False
Call LoadCmb

Call Display_Value
Call ST_UPD_MODE(False)
Call INI_Controls(Me)
EESNameSort = True
Screen.MousePointer = DEFAULT

lblOHire.Caption = lStr("Original Hire")
lblPT.Caption = lStr("Category")
lblTitle(24).Caption = lStr("Region")
lblTitle(23).Caption = lStr("Location")

cmdRecal.Enabled = gSec_Upd_AffirmAction_Data

' danielk - 12/31/2002 - added next line, w/o it we tried to SetFocus when the form was hidden
Me.Show
'cmdNew.SetFocus

End Sub
Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm 'Ticket #30482 Franks 08/16/2017
End Sub

Private Function LoadCmb()
Dim ctylist, x

cmbType.AddItem "Employee"
cmbType.AddItem "Applicant"

cmbDisability.AddItem "Yes"
cmbDisability.AddItem "No"

cmbVietnam.AddItem "Yes"
cmbVietnam.AddItem "No"

cmbVeteran.AddItem "Yes"
cmbVeteran.AddItem "No"

'Ticket #22682 - Release 8.0 - Ethnicity
cmbEthnicity.AddItem "Hispanic or Latino"
cmbEthnicity.AddItem "Not Hispanic or Latino"

'Call function to populate the dropdown list with Countries from MTF file
ctylist = CountryList
x = 1
Do While x > 0
    x = InStr(ctylist, "&")
    If x > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, x - 1)
        'comCountryOfEmpOfEmp.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, x + 1)
    Else
        comCountryOfEmp.AddItem ctylist
        'comCountryOfEmpOfEmp.AddItem ctylist
    End If
Loop
comCountryOfEmp.ListIndex = -1

End Function

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

If InStr(xCountryList, comCountryOfEmp) = 0 And comCountryOfEmp <> "" Then
    xCountryList = xCountryList & "&" & comCountryOfEmp
    comCountryOfEmp.AddItem comCountryOfEmp
End If

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
    Resume Next
End If
End Function


Private Function chkEEO()
Dim rsTA As New ADODB.Recordset

chkEEO = False

If xActn = "A" Then

    rsTA.Open "SELECT EO_EEONNBR FROM HREEO WHERE EO_EEONNBR = " & Val(txtEEO), gdbAdoIhr001, adOpenStatic
    If Not rsTA.EOF Then
        MsgBox "EEO Number Already Exists!"
        txtEEO.SetFocus
        Exit Function
    End If
    rsTA.Close
    If cmbType.ListIndex = 0 Then
        OETYPE = "E"
    Else
        OETYPE = "A"
    End If
    If OETYPE = "E" Then
        If Len(elpEEID.Text) = 0 Then
            MsgBox "Employee Number is a required field!"
            elpEEID.SetFocus
            Exit Function
        End If
    End If
    If Len(elpEEID.Text) > 0 Then
        If elpEEID.Caption = "Unassigned" Then
            MsgBox "Invalid Employee Number!"
            elpEEID.SetFocus
            Exit Function
        End If
        txtEEO = elpEEID.Text
    End If
End If
    
If Len(txtSurname) < 1 Then
    MsgBox "Surname is a required field"
    txtSurname.SetFocus
    Exit Function
End If
If Len(txtFName) < 1 Then
    MsgBox lStr("First Name is a required field")
    txtFName.SetFocus
    Exit Function
End If

If gSec_Show_DOB Then 'Ticket #23947 Franks 06/20/2013
    ' dkostka - 02/27/01 - DOB is NOT required.
    If Len(dlpDOB.Text) < 1 Then
        MsgBox "Birth Date is a required field"
        dlpDOB.SetFocus
        Exit Function
    Else
        If Not IsDate(dlpDOB.Text) And Format(dlpDOB.Text, "@") <> "" Then
            MsgBox "Invalid Birthdate"
            dlpDOB.SetFocus
            Exit Function
        End If
    End If
End If

If Trim(txtGender) = "" Then
    MsgBox "Gender is a required field"
    optSex(0) = True: txtGender = "M"
    optSex(0).SetFocus
    Exit Function
End If

'Ticket #22682 - Release 8.0 - Ethnicity is mandatory
If cmbEthnicity.ListIndex = -1 Then
    MsgBox "Ethnicity is a required field"
    cmbEthnicity.SetFocus
    Exit Function
End If

If clpCode(1).Text = "" Then
    MsgBox "For Race you must enter a valid Code!!"
    clpCode(1).SetFocus
    Exit Function
End If

'Ticket #22682 - Release 8.0
If Len(clpCode(4).Text) > 0 Then
    If Not clpCode(4).ListChecker Then
        clpCode(4).SetFocus
        Exit Function
    End If
End If

'Ticket #18790 - begin
If Len(dlpDOH) < 1 Then
    MsgBox lblOHire.Caption & " is a required field."
    dlpDOH.SetFocus
    Exit Function
Else
    If Not IsDate(dlpDOH.Text) And Format(dlpDOH.Text, "@") <> "" Then
        MsgBox "Invalid " & lblOHire.Caption
        dlpDOH.SetFocus
        Exit Function
    End If
End If
If Len(comCountryOfEmp.Text) < 1 Then
    MsgBox "Country of Employment is a required field."
    comCountryOfEmp.SetFocus
    Exit Function
End If
'Ticket #18790 - end

If clpNOGC.Text = "" Then
    MsgBox "EEO Job Category is a required field!"
    clpNOGC.Text = ""
    clpNOGC.SetFocus
    Exit Function
End If

If cmbDisability.ListIndex = 0 And clpCode(2).Text = "" Then
    MsgBox "Disability is a required field"
    clpCode(2).SetFocus
    Exit Function
Else
    If clpCode(2).Text <> "" Then
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "Disability code must be valid"
             clpCode(2).SetFocus
            Exit Function
        End If
    End If
End If

If clpNOGC.Text <> "" Then
    If clpNOGC.Caption = "Unassigned" Then
        MsgBox "N.O.C. code must be valid"
         clpNOGC.Text = ""
         clpNOGC.SetFocus
        Exit Function
    End If
End If

If cmbVeteran.ListIndex = 1 Then
    cmbVietnam.ListIndex = 1
End If

chkEEO = True

End Function
Sub NOC_Desc()
Dim SQLQ As String

On Error GoTo NOCd_Err

 clpNOGC.Caption = "Unassigned"
If Len(clpNOGC.Text) > 0 Then
    SQLQ = "OC_CODE = '" & clpNOGC.Text & "'"
    NOC_Snap.Requery
    NOC_Snap.Find SQLQ
    If Not NOC_Snap.EOF Then
         clpNOGC.Caption = NOC_Snap("OC_SDESCR")
         clpNOGC.ShowDescription = True
    End If
End If

Exit Sub

NOCd_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "NOC Snap", "NOC Code", "SELECT")
Call RollBack

End Sub

Private Sub OETYPE_Change()

If Len(OETYPE) > 0 Then
    If OETYPE.Text = "E" Then
        cmbType.ListIndex = 0
    Else
        cmbType.ListIndex = 1
    End If
End If

End Sub

Sub optSex_Click(Index As Integer, Value As Integer)

End Sub

Sub Plan_Desc()
Dim SQLQ As String

Exit Sub

PlanNo_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Plan Snap", "Plan Number", "SELECT")
Call RollBack

End Sub


Sub txtAborig_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub optSex_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If optSex(0) = True Then txtGender = "M" Else txtGender = "F"
End Sub

Private Sub txtCountryOfEmp_Change()
comCountryOfEmp.Text = txtCountryOfEmp.Text
End Sub

Sub txtEEO_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub txtEEOJobCat_DblClick()
    ' clpNOGC_DblClick
End Sub

Sub txtEEOJobCat_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtEESearch_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEESearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub txtFName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtGender_Change()
If Len(Trim(txtGender)) > 0 Then
    If txtGender = "M" Then
        optSex(0) = True
    Else
        optSex(1) = True
    End If
End If
End Sub

Sub txtGender_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtSIN_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSurname_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If xActTermFlag = "ACT" Then
            SQLQ = "SELECT * FROM HREEO "
            SQLQ = SQLQ & "WHERE EO_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ") "
            SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        End If
        If xActTermFlag = "TERM" Then
            SQLQ = "SELECT * FROM Term_HREEO "
            SQLQ = SQLQ & "WHERE EO_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn & ") "
            SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        End If
        
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim SQLQ As String
Dim countr   As Integer

Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
''' Comment by Frank May 30,2002 for Delete button error
''' Before delete a record, caused problem
Call Display_Value
End Sub

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


Private Sub Race_Desc()
Dim SQLQ As String

On Error GoTo RaceDesc_Err

 clpCode(1).Caption = "Unassigned"
If Len(clpCode(1).Text) > 0 Then
    SQLQ = "TB_KEY = '" & clpCode(1).Text & "'"
    Race_Snap.Requery
    Race_Snap.Find SQLQ
    If Not Race_Snap.EOF Then
         clpCode(1).Caption = Race_Snap("TB_DESC")
         clpCode(1).ShowDescription = True
    End If
End If
Exit Sub

RaceDesc_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Race Snap", "Race Code", "SELECT")
Call RollBack

End Sub

Private Sub CR_Race_Snap()

Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo Race_Err
Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HRTABL"

If Race_Snap.State <> 0 Then Race_Snap.Close
Race_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

If Race_Snap.EOF And Race_Snap.BOF Then
    Msg = "No Race descriptions found."
    MsgBox Msg
    Exit Sub
Else
  Race_Snap.MoveFirst
End If

Screen.MousePointer = DEFAULT

Exit Sub

Race_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Race", "HRTABL", "SELECT")
Call RollBack

End Sub




Private Sub CR_HREMP_SNAP()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo HR_Err
Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HREMP"

If HR_Snap.State <> 0 Then HR_Snap.Close
HR_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

If HR_Snap.EOF And HR_Snap.BOF Then
    Msg = "No Records found."
    MsgBox Msg
    Exit Sub
Else
    HR_Snap.MoveFirst
End If

Screen.MousePointer = DEFAULT

Exit Sub

HR_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "", "FROM HREMP", "SELECT")
Call RollBack

End Sub


Private Sub Veteran_Change()
    cmbVeteran.Text = IIf(Veteran.Text = "0", "No", "Yes")
End Sub

Private Sub Vietnam_Change()
    cmbVietnam.Text = IIf(Vietnam.Text = "0", "No", "Yes")
End Sub


''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close

            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

        Exit Sub
    End If
    
    If xActTermFlag = "ACT" Then
        SQLQ = "Select HREEO.* "
        SQLQ = SQLQ & " FROM HREEO"
        SQLQ = SQLQ & " WHERE EO_EEONNBR = " & Data1.Recordset!EO_EEONNBR
    End If
    If xActTermFlag = "TERM" Then
        SQLQ = "Select Term_HREEO.*, TERM_DOT,SH_SALARY "
        SQLQ = SQLQ & " FROM Term_HREEO "
        SQLQ = SQLQ & "LEFT JOIN Term_HRTRMEMP ON Term_HREEO.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ "
        SQLQ = SQLQ & "LEFT JOIN Term_SALARY_HISTORY ON Term_HREEO.TERM_SEQ=Term_SALARY_HISTORY.TERM_SEQ "
        SQLQ = SQLQ & " WHERE EO_EEONNBR = " & Data1.Recordset!EO_EEONNBR
        SQLQ = SQLQ & " AND NOT Term_SALARY_HISTORY.SH_CURRENT = 0 "
    End If
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic


    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    elpEEID.Enabled = False
    txtEEO.Enabled = False

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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = True
End Property

Public Property Get Addable() As Boolean

Addable = gSec_Upd_AffirmAction_Data  'True
End Property
Public Property Get Updateble() As Boolean
Updateble = gSec_Upd_AffirmAction_Data  'True
End Property
Public Property Get Deleteble() As Boolean

Deleteble = gSec_Upd_AffirmAction_Data  'True
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
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
'TF = False
Call ST_UPD_MODE(TF)
End Sub

