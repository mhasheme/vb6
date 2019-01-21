VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSDept 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Department Security"
   ClientHeight    =   9450
   ClientLeft      =   465
   ClientTop       =   1410
   ClientWidth     =   11220
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9450
   ScaleWidth      =   11220
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PD_LDATE"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   7740
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PD_LTIME"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PD_LUSER"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   9960
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   900
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11220
      _Version        =   65536
      _ExtentX        =   19791
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
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   135
         Width           =   660
      End
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USERID"
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
         TabIndex        =   24
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
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
         Left            =   3030
         TabIndex        =   23
         Top             =   120
         Width           =   1290
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   28
      Top             =   8790
      Width           =   11220
      _Version        =   65536
      _ExtentX        =   19791
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
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   870
         TabIndex        =   16
         Tag             =   "Edit the information "
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   30
         TabIndex        =   17
         Tag             =   "Close and exit this screen"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1725
         TabIndex        =   12
         Tag             =   "Save the changes made"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2565
         TabIndex        =   13
         Tag             =   "Cancel the changes made"
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3450
         TabIndex        =   14
         Tag             =   "Add a new Record"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4305
         TabIndex        =   15
         Tag             =   "Delete the Record Selected"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5145
         TabIndex        =   18
         Tag             =   "Print Listing "
         Top             =   180
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6240
         Top             =   120
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
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   7560
         Top             =   30
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   7560
         Top             =   390
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsdept.frx":0000
      Height          =   3285
      Left            =   240
      OleObjectBlob   =   "fsdept.frx":0014
      TabIndex        =   0
      Top             =   510
      Width           =   10770
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "PD_DEPT"
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Tag             =   "01-Department - Code"
      Top             =   3900
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "PD_DIV"
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Tag             =   "01-Division - Code"
      Top             =   4604
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PD_ORG"
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Tag             =   "01-Union - Code"
      Top             =   4252
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PD_SECTION"
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Tag             =   "01-Section - Code"
      Top             =   4956
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PD_ADMINBY"
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Tag             =   "01-Administered By - Code"
      Top             =   5308
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEIDIncl 
      DataField       =   "PD_INCLEMPNBR"
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Tag             =   "10-Enter Employee Number to Include"
      ToolTipText     =   "List of employees to include who are not otherwise part of any of the groups in the table above."
      Top             =   7800
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEIDExcl 
      DataField       =   "PD_EXCLEMPNBR"
      Height          =   285
      Left            =   2640
      TabIndex        =   10
      Tag             =   "10-Enter Employee Number to Exclude"
      Top             =   7080
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PD_LOC"
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   6
      Tag             =   "01-Section - Code"
      Top             =   5660
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PD_REGION"
      Height          =   285
      Index           =   4
      Left            =   2640
      TabIndex        =   7
      Tag             =   "01-Administered By - Code"
      Top             =   6012
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PD_SUPCODE"
      Height          =   285
      Index           =   5
      Left            =   2640
      TabIndex        =   8
      Tag             =   "00-Supervisory Code for cheque sorting "
      Top             =   6364
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSP"
   End
   Begin INFOHR_Controls.CodeLookup clpVadim2 
      DataField       =   "PD_VADIM2"
      Height          =   285
      Left            =   2640
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDV2"
   End
   Begin VB.Label lblVadim2 
      AutoSize        =   -1  'True
      Caption         =   "Vadim Field 2"
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
      Left            =   450
      TabIndex        =   40
      Top             =   6765
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblSupervisor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Supervisor Code"
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
      Left            =   450
      TabIndex        =   39
      Top             =   6409
      Visible         =   0   'False
      Width           =   1410
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
      Left            =   450
      TabIndex        =   38
      Top             =   5705
      Width           =   615
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
      Left            =   450
      TabIndex        =   37
      Top             =   6057
      Width           =   510
   End
   Begin VB.Label lblhelptxt2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "List employees to include who are not otherwise part of any of the groups in the table above."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3000
      TabIndex        =   36
      Top             =   8400
      Width           =   7920
   End
   Begin VB.Label lblhelptxt1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "You must choose one of the rows in the table above."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3000
      TabIndex        =   35
      Top             =   8160
      Width           =   4515
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Employee Numbers"
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
      Left            =   450
      TabIndex        =   34
      Top             =   7125
      Width           =   1980
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Include Employee Numbers"
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
      Index           =   0
      Left            =   480
      TabIndex        =   33
      Top             =   7845
      Width           =   1935
   End
   Begin VB.Label lblAdminBy 
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
      Left            =   450
      TabIndex        =   32
      Top             =   5353
      Width           =   1365
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
      Left            =   450
      TabIndex        =   31
      Top             =   4649
      Width           =   555
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
      Left            =   450
      TabIndex        =   30
      Top             =   5001
      Width           =   540
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
      Left            =   450
      TabIndex        =   29
      Top             =   4297
      Width           =   960
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   450
      TabIndex        =   27
      Top             =   3945
      Width           =   990
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "PD_COMPNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   10080
      TabIndex        =   26
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_Return 
         Caption         =   "&Return to Security"
      End
   End
End
Attribute VB_Name = "frmSDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDATA As New ADODB.Recordset

Private Function chkSecDept()

Dim SQLQ As String, Msg As String, dd#, PID&, Dept$

chkSecDept = False

On Error GoTo chkSecDept_Err

'Hemu 05/13/2003 Begin
If Len(clpDept.Text) = 0 Then
    MsgBox "Department Code is required"
    clpDept.SetFocus
    Exit Function
End If

If clpDept.Caption = "Unassigned" Then
    MsgBox "Department Code must be valid"
    clpDept.SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
    If clpCode(1).Caption = "Unassigned" And Left(clpCode(1).Text, 1) <> "-" Then
        MsgBox lblUnion.Caption & " must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
Else
    If clpCode(1).Caption = "Unassigned" And clpCode(1).Text <> "-NON" And clpCode(1).Text <> "-EXE" Then
        MsgBox lblUnion.Caption & " must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If clpDiv.Caption = "Unassigned" Then
    MsgBox lblDiv.Caption & " must be valid"
    clpDiv.SetFocus
    Exit Function
End If
If clpCode(2).Caption = "Unassigned" Then
    MsgBox lblSection.Caption & " must be valid"
    clpCode(2).SetFocus
    Exit Function
End If
'Hemu 05/13/2003 End

'Ticket #18235
If clpCode(0).Caption = "Unassigned" Then
    MsgBox lblAdminBy.Caption & " must be valid"
    clpCode(0).SetFocus
    Exit Function
End If

'Ticket #22682 - Release 8.0
If clpCode(3).Caption = "Unassigned" Then
    MsgBox lblLocation.Caption & " must be valid"
    clpCode(3).SetFocus
    Exit Function
End If

'Ticket #22682 - Release 8.0
If clpCode(4).Caption = "Unassigned" Then
    MsgBox lblRegion.Caption & " must be valid"
    clpCode(4).SetFocus
    Exit Function
End If

'Ticket #24161 - Samuel Only - Release 8.0
If glbSamuel Then
    If clpCode(5).Caption = "Unassigned" Then
        MsgBox lblSupervisor.Caption & " must be valid"
        clpCode(5).SetFocus
        Exit Function
    End If
    If clpVadim2.Caption = "Unassigned" Then
        MsgBox lblVadim2.Caption & " must be valid"
        clpVadim2.SetFocus
        Exit Function
    End If
End If

If Not elpEEIDIncl.ListChecker Then
    Exit Function
End If

If Not elpEEIDExcl.ListChecker Then
    Exit Function
End If

If Len(elpEEIDIncl.Text) > 500 Then
    MsgBox "Include Employee Numbers cannot exceed 500 characters"
    elpEEIDIncl.SetFocus
    Exit Function
End If

If Len(elpEEIDExcl.Text) > 500 Then
    MsgBox "Exclude Employee Numbers cannot exceed 500 characters"
    elpEEIDIncl.SetFocus
    Exit Function
End If

If IsNull(rsDATA!PD_ID) Then PID& = 0 Else PID& = Val(rsDATA!PD_ID)
Dept$ = clpDept.Text
'Ticket #24161 - Samuel Only - Release 8.0
If glbSamuel Then
    If modISDupDept(glbSecUSERID, Dept$, clpCode(1).Text, clpDiv.Text, clpCode(2).Text, PID&, clpCode(0).Text, clpCode(3).Text, clpCode(4).Text, elpEEIDIncl.Text, elpEEIDExcl.Text, clpCode(5).Text, clpVadim2.Text) Then
        MsgBox lStr("Department, Union, Division, Section, Administered By, Location, Region, Supervisor Code, Vadim Field 2 and Include Employee Nos., Exclude Employee Nos. must be unique")
        clpDept.SetFocus
        Exit Function
    End If
Else
    If modISDupDept(glbSecUSERID, Dept$, clpCode(1).Text, clpDiv.Text, clpCode(2).Text, PID&, clpCode(0).Text, clpCode(3).Text, clpCode(4).Text, elpEEIDIncl.Text, elpEEIDExcl.Text) Then
        MsgBox lStr("Department, Union, Division, Section, Administered By, Location, Region and Include Employee Nos., Exclude Employee Nos. must be unique")
        clpDept.SetFocus
        Exit Function
    End If
End If

chkSecDept = True

Exit Function

chkSecDept_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSecDept", "HRPASDEP", "validation")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub cmdCancel_Click()

On Error GoTo Can_Err

rsDATA.CancelUpdate

Call Display_Value

Call ST_UPD_MODE(False)  ' reset screen's attributes

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRPASDEP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdCancel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "


a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HRPASDEP", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdDelete_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

clpDept.SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRPASDEP", "Modify")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdNew_Click()
Dim SQLQ As String

'Ticket #18668 - Remove the limit - esp. for Linamar
'I tested the WHERE clause of the Dept security and it worked fine and it also worked for reports
'as well - I tried to setup 71 depts security rows and it worked fine. Also searched internet -
'nowhere I could find any limit to WHERE clause.
'So we are not sure why this was added
'If Data1.Recordset.RecordCount = 50 Then
'    MsgBox "You can't add more than 50 departments"
'    Exit Sub
'End If

Call ST_UPD_MODE(True)

On Error GoTo AddN_Err

Call Set_Control("B", Me)

rsDATA.AddNew

clpDept.Caption = ""
lblCNum.Caption = "001"
lblUSERID.Caption = glbSecUSERID

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRPASDEP", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub CmdNew_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim x%
Dim xID
Dim xTemplate As String

On Error GoTo OK_Err

'Ticket #21629 - Jerry said Department Security of the user is independent of their Template

''Ticket #20585 - If Template then update users with this template as well.
''If User and with no template then update that user's profile.
''if User and with Template then do not update user's profile.
''Get the Template Name of this User ID
'xTemplate = Get_Template(glbSecUSERID)
'
'If xTemplate = "TEMPLATE" Then
'    'Update all users with this template. After the changes are saved
'ElseIf xTemplate = "" Then
'    'User - User with no template - don't do anything let system update user's profile
'ElseIf xTemplate <> "TEMPLATE" Then
'    'User with template - do not allow to save these changes.
'    MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
'End If
'
''if Template or User
'If xTemplate = "TEMPLATE" Or xTemplate = "" Then

    If Not chkSecDept() Then Exit Sub
    
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    
    Call Set_Control("U", Me, rsDATA)
    
    rsDATA("PD_USERID") = lblUSERID & ""
    
    If glbLinamar Then
        If clpCode(2).Text <> "" And IsNull(clpCode(2).Text) = False Then
            rsDATA("PD_SECTION") = clpDiv & clpCode(2)
        Else
            rsDATA("PD_SECTION") = ""
        End If
    End If
    
    rsDATA("PD_INCLEMPNBR") = elpEEIDIncl.Text
    rsDATA("PD_EXCLEMPNBR") = elpEEIDExcl.Text
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
    
    gdbAdoIhr001.Execute "UPDATE HRPASDEP SET PD_INCLEMPNBR = '" & elpEEIDIncl.Text & "' WHERE PD_USERID ='" & Replace(lblUSERID, "'", "''") & "'"
'End If

Data1.Refresh

'Ticket #21629 - Jerry said not to change user's Dept security based on Template's Dept Security
'Ticket #20585 - Security Based on Template Profile
'If xTemplate = "TEMPLATE" Then
'    'Call procedure to Update all users with this template.
'    Call Update_Users_withthis_Template(glbSecUSERID)
'End If

Call ST_UPD_MODE(False)

Me.vbxTrueGrid.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPASDEP", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName.Caption & "'s Departmental Security"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRPASDEP", "SELECT")

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
Dim xTemplate  As String

Screen.MousePointer = HOURGLASS
lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName

'Ticket #29024 - Had to comment this so this form can be shown Modal form from frmSecurity
'frmSDept.Show

Me.Caption = lStr("Department Security - ") & lblEEName


Data1.ConnectionString = glbAdoIHRDB

x% = modGetDepts()

Call setCaption(lblDiv)
Call setCaption(lblDept)
Call setCaption(lblSection)
Call setCaption(lblUnion)
Call setCaption(lblAdminBy)
'Ticket #22682 - Release 8.0
Call setCaption(lblLocation)
Call setCaption(lblRegion)

Call setCaption(Me.vbxTrueGrid.Columns(0))
Call setCaption(Me.vbxTrueGrid.Columns(1))
Call setCaption(Me.vbxTrueGrid.Columns(2))
Call setCaption(Me.vbxTrueGrid.Columns(3))
Call setCaption(Me.vbxTrueGrid.Columns(4))
'Ticket #22682 - Release 8.0
Call setCaption(Me.vbxTrueGrid.Columns(5))
Call setCaption(Me.vbxTrueGrid.Columns(6))

'Ticket #24161 - Samuel Only - Release 8.0
If glbSamuel Then
    Call setCaption(lblSupervisor)
    Call setCaption(lblVadim2)
    Call setCaption(Me.vbxTrueGrid.Columns(7))
    Call setCaption(Me.vbxTrueGrid.Columns(8))
Else
    Me.vbxTrueGrid.Columns(7).Visible = False
    Me.vbxTrueGrid.Columns(8).Visible = False
End If

If vbxTrueGrid.Visible Then
    Me.vbxTrueGrid.Columns(8).Visible = False
    Me.vbxTrueGrid.SetFocus
End If

'Ticket #24161 - Samuel Only - Release 8.0
If glbSamuel Then
    lblSupervisor.Visible = True
    clpCode(5).Visible = True
    lblVadim2.Visible = True
    clpVadim2.Visible = True
Else
    lblSupervisor.Visible = False
    clpCode(5).Visible = False
    lblVadim2.Visible = False
    clpVadim2.Visible = False
End If

Call INI_Controls(Me)

Call ST_UPD_MODE(False)

'Ticket #21629 - Jerry said Department Security is independent of Template Security
'Ticket #20585 - Enable/Disable Edit, New and Delete buttons based on the type of user
'xTemplate = Get_Template(glbSecUSERID)
'If xTemplate = "" Or xTemplate = "TEMPLATE" Then
'    'User without Template or Template
'Else
'    'User with Template
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmSDept = Nothing
End Sub

Private Sub mnu_File_Exit_Click()
    Call ApplicationEnd
End Sub

Private Sub mnu_F_PrintSetup_Click()
MDIMain.vbxCommon.Action = 5
End Sub


Private Sub mnu_Return_Click()
   Unload Me
End Sub

Private Function modGetDepts()
Dim SQLQ

modGetDepts = False

Screen.MousePointer = HOURGLASS

On Error GoTo modGetDeptsErr

    SQLQ = "SELECT PD_ID,PD_USERID,PD_DEPT,PD_ORG_TABL,PD_ORG,PD_DIV,PD_SECTION,PD_SECTION_TABL,"
    SQLQ = SQLQ & "PD_LDATE,PD_LTIME,PD_LUSER,PD_ADMINBY,PD_ADMINBY_TABL,PD_LOC,PD_LOC_TABL,"
    SQLQ = SQLQ & "PD_REGION,PD_REGION_TABL,PD_INCLEMPNBR,PD_EXCLEMPNBR "

'Ticket #24161 - Samuel Only - Release 8.0
If glbSamuel Then
    SQLQ = SQLQ & ",PD_SUPCODE,PD_SUPCODE_TABL,PD_VADIM2,PD_VADIM2_TABL"
End If

If glbLinamar Then
    SQLQ = SQLQ & ", SUBSTRING(PD_SECTION,4,20) AS SHOWSECTION from HRPASDEP " ' saved query object
Else
    SQLQ = SQLQ & ",PD_SECTION AS SHOWSECTION from HRPASDEP "  ' saved query object
End If
SQLQ = SQLQ & " WHERE PD_USERID = '" & Replace(glbSecUSERID, "'", "''") & "'"
SQLQ = SQLQ & " ORDER BY PD_DEPT"

Data1.RecordSource = SQLQ
Data1.Refresh


modGetDepts = True
Screen.MousePointer = DEFAULT

Exit Function

modGetDeptsErr:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Security Depts", "HRPASDEP", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Exit Function

End Function

Private Function modISDupDept(UserID As String, Dept As String, Org As String, Div As String, SECTION As String, PID As Long, AdminBy As String, Loc As String, Region As String, IncludeEmp As String, ExcludeEmp As String, Optional SupCode As String, Optional Vadim2 As String)
Dim SQLQ As String
Dim snapDepSec As New ADODB.Recordset

modISDupDept = True

On Error GoTo modISDupDept_Err

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HRPASDEP "
SQLQ = SQLQ & "Where PD_USERID = '" & Replace(UserID, "'", "''") & "'"
SQLQ = SQLQ & " AND PD_DEPT = '" & Dept & "' "
SQLQ = SQLQ & " AND PD_ID <> " & PID
If Len(Org) > 0 Then
    SQLQ = SQLQ & " AND PD_ORG = '" & Org & "' "
Else
    SQLQ = SQLQ & " AND PD_ORG IS NULL "
End If
If Len(Div) > 0 Then
    SQLQ = SQLQ & " AND PD_DIV = '" & Div & "' "
Else
    SQLQ = SQLQ & " AND PD_DIV IS NULL "
End If
If Len(SECTION) > 0 Then
    If glbLinamar Then
        SQLQ = SQLQ & " AND PD_SECTION = '" & clpDiv & SECTION & "' "
    Else
        SQLQ = SQLQ & " AND PD_SECTION = '" & SECTION & "' "
    End If
Else
    SQLQ = SQLQ & " AND PD_SECTION IS NULL "
End If
'Ticket #18235
If Len(AdminBy) > 0 Then
    SQLQ = SQLQ & " AND PD_ADMINBY = '" & AdminBy & "' "
Else
    SQLQ = SQLQ & " AND PD_ADMINBY IS NULL "
End If

'Ticket #22682 - Release 8.0
If Len(Loc) > 0 Then
    SQLQ = SQLQ & " AND PD_LOC = '" & Loc & "' "
Else
    SQLQ = SQLQ & " AND PD_LOC IS NULL "
End If

'Ticket #22682 - Release 8.0
If Len(Region) > 0 Then
    SQLQ = SQLQ & " AND PD_REGION = '" & Region & "' "
Else
    SQLQ = SQLQ & " AND PD_REGION IS NULL "
End If

'Ticket #24161 - Samuel Only - Release 8.0
If glbSamuel Then
    If Len(SupCode) > 0 Then
        SQLQ = SQLQ & " AND PD_SUPCODE = '" & SupCode & "' "
    Else
        SQLQ = SQLQ & " AND PD_SUPCODE IS NULL "
    End If
    If Len(Vadim2) > 0 Then
        SQLQ = SQLQ & " AND PD_VADIM2 = '" & Vadim2 & "' "
    Else
        SQLQ = SQLQ & " AND PD_VADIM2 IS NULL "
    End If
End If

'7.9 - Enhancement
If Len(IncludeEmp) > 0 Then
    SQLQ = SQLQ & " AND PD_INCLEMPNBR LIKE '" & getEmpnbr(IncludeEmp) & "' "
'Else
'    SQLQ = SQLQ & " AND PD_INCLEMPNBR IS NULL "
End If
If Len(ExcludeEmp) > 0 Then
    SQLQ = SQLQ & " AND PD_EXCLEMPNBR LIKE '" & getEmpnbr(ExcludeEmp) & "' "
'Else
'    SQLQ = SQLQ & " AND PD_EXCLEMPNBR IS NULL "
End If

snapDepSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapDepSec.BOF And snapDepSec.EOF Then
    modISDupDept = False
End If

snapDepSec.Close
Screen.MousePointer = DEFAULT


Exit Function

modISDupDept_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Find Duplicate", "HRPASDEP", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

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
    
    glbOHSEdit% = TF
    
    cmdOK.Enabled = TF
    cmdCancel.Enabled = TF
    
    cmdClose.Enabled = FT
    cmdNew.Enabled = FT
    cmdModify.Enabled = FT
    cmdDelete.Enabled = FT
    cmdPrint.Enabled = FT
    clpDept.Enabled = TF
    clpCode(1).Enabled = TF
    clpCode(2).Enabled = TF     'Lucy July 6, 2000
    clpDiv.Enabled = TF
    vbxTrueGrid.Enabled = FT
    clpCode(0).Enabled = TF     'Administered By
    clpCode(3).Enabled = TF     'Location
    clpCode(4).Enabled = TF     'Region
    elpEEIDIncl.Enabled = TF
    elpEEIDExcl.Enabled = TF
    
    'Ticket #24161 - Samuel Only - Release 8.0
    If glbSamuel Then
        clpCode(5).Enabled = TF
        clpVadim2.Enabled = TF
    End If
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        cmdModify.Enabled = False
    End If
    
    If Not gSec_Upd_Security Then 'And Not gSec_Upd_Quick_ESS Then
        cmdModify.Enabled = False
        cmdNew.Enabled = False
        cmdDelete.Enabled = False
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
    
    SQLQ = "SELECT PD_ID,PD_USERID,PD_DEPT,PD_ORG_TABL,PD_ORG,PD_DIV,PD_SECTION,PD_SECTION_TABL,"
    SQLQ = SQLQ & "PD_LDATE,PD_LTIME,PD_LUSER,PD_ADMINBY,PD_ADMINBY_TABL,PD_LOC,PD_LOC_TABL,"
    SQLQ = SQLQ & "PD_REGION,PD_REGION_TABL,PD_INCLEMPNBR,PD_EXCLEMPNBR "
    
    'Ticket #24161 - Samuel Only - Release 8.0
    If glbSamuel Then
        SQLQ = SQLQ & ",PD_SUPCODE,PD_SUPCODE_TABL,PD_VADIM2,PD_VADIM2_TABL"
    End If
    
    If glbLinamar Then
        SQLQ = SQLQ & ", SUBSTRING(PD_SECTION,4,20) AS SHOWSECTION from HRPASDEP " ' saved query object
    Else
        SQLQ = SQLQ & ",PD_SECTION AS SHOWSECTION from HRPASDEP "  ' saved query object
    End If
    SQLQ = SQLQ & " WHERE PD_USERID = '" & Replace(glbSecUSERID, "'", "''") & "'"
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then ' if the tab key was struck
        KeyAscii = 0
        If cmdOK.Enabled Then
            cmdOK.SetFocus
        Else
            cmdClose.SetFocus
        End If
    End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, x%

Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value
Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRJOBEVL", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    
    SQLQ = "SELECT * from HRPASDEP "
    SQLQ = SQLQ & " WHERE PD_ID = " & Data1.Recordset!PD_ID

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
         
    If glbLinamar Then
        clpCode(2) = Mid(rsDATA("PD_SECTION") & "", 4)
    End If

End Sub

Private Sub Update_Users_withthis_Template(xTemplate)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Retrieve all users associated with this changed Template
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & xTemplate & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            'Update each user with this changed Template
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "DEPARTMENT")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

