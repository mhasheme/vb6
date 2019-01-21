VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEHSCost 
   AutoRedraw      =   -1  'True
   Caption         =   "Accident Cost Analysis"
   ClientHeight    =   8595
   ClientLeft      =   -135
   ClientTop       =   600
   ClientWidth     =   11460
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fehscost.frx":0000
      Height          =   1215
      Left            =   120
      OleObjectBlob   =   "Fehscost.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   10935
   End
   Begin MSMask.MaskEdBox medUnemp 
      DataField       =   "AC_Unemp"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   42
      Tag             =   "00-Unemployment and disablity costs"
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCompe 
      DataField       =   "AC_Compe"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   41
      Tag             =   "00-Workers compensation (actual costs)"
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   11040
      Top             =   8280
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
      Caption         =   "Ado3"
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   11040
      Top             =   8040
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
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      DataField       =   "AC_Case"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   4830
      MaxLength       =   8
      TabIndex        =   2
      Tag             =   "11- incident Number"
      Top             =   1830
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox comShift 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Tag             =   "01-Incident Number"
      Top             =   1830
      Width           =   1575
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   7935
      Width           =   11460
      _Version        =   65536
      _ExtentX        =   20214
      _ExtentY        =   1164
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   11730
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
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox Updstats 
      DataField       =   "AC_LDate"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   6150
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "AC_LTime"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7950
      MaxLength       =   25
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "AC_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   9630
      MaxLength       =   25
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSMask.MaskEdBox medVehic 
      DataField       =   "AC_Vehic"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   43
      Tag             =   "00-Damage costs - Vehicle"
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPrope 
      DataField       =   "AC_Prope"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   44
      Tag             =   "00-Damage costs - Property"
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11460
      _Version        =   65536
      _ExtentX        =   20214
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEENumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
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
         TabIndex        =   8
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
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
         TabIndex        =   7
         Top             =   135
         Width           =   720
      End
   End
   Begin MSMask.MaskEdBox medFirst 
      DataField       =   "AC_First"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   45
      Tag             =   "00-First aid costs"
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medMater 
      DataField       =   "AC_Mater"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   46
      Tag             =   "00-Material lost"
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHourly 
      DataField       =   "AC_Hour"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   47
      Tag             =   "00-Approximated hourly rate"
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPeopl 
      DataField       =   "AC_Peopl"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   48
      Tag             =   "00-# of people involved in investigation"
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTotDC 
      Height          =   285
      Left            =   7920
      TabIndex        =   49
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medSick 
      DataField       =   "AC_Sick"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   50
      Tag             =   "00-Sickness and accident costs"
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medDispo 
      DataField       =   "AC_Displ"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   51
      Tag             =   "00-Disposal costs"
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medClean 
      DataField       =   "AC_Clean"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   52
      Tag             =   "00-Clean-up costs"
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTime 
      DataField       =   "AC_Time"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   53
      Tag             =   "00-Time spent on investigation (hours)"
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOTime 
      DataField       =   "AC_OTime"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   54
      Tag             =   "00-Overtime"
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTotal 
      Height          =   285
      Left            =   3240
      TabIndex        =   55
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAmount 
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3240
      TabIndex        =   56
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medRent 
      DataField       =   "AC_Rent"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   57
      Tag             =   "00-Rentals"
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medSubco 
      DataField       =   "AC_Subco"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   58
      Tag             =   "00-Sub-contracting costs"
      Top             =   4440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTrans 
      DataField       =   "AC_Trans"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   59
      Tag             =   "00-Transportation costs"
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTrain 
      DataField       =   "AC_Train"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   60
      Tag             =   "00-Training costs"
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medDown 
      DataField       =   "AC_Down"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   61
      Tag             =   "00-Downtime (production loss)"
      Top             =   5880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTotIC 
      Height          =   285
      Left            =   7920
      TabIndex        =   62
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medCust 
      DataField       =   "AC_Custo"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   7920
      TabIndex        =   63
      Tag             =   "00-Customer relations && administration"
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblUserDesc 
      Height          =   255
      Left            =   4200
      TabIndex        =   67
      Top             =   6960
      Width           =   2415
   End
   Begin VB.Label lblUpdateBy 
      Caption         =   "Updated By"
      Height          =   255
      Left            =   3240
      TabIndex        =   66
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label lblUpdDateDesc 
      Height          =   255
      Left            =   7680
      TabIndex        =   65
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Label lblUpdateDate 
      Caption         =   "Updated Date"
      Height          =   255
      Left            =   6600
      TabIndex        =   64
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Caption         =   "Total Cost (IC)"
      Height          =   255
      Index           =   16
      Left            =   6360
      TabIndex        =   40
      Top             =   6600
      Width           =   1305
   End
   Begin VB.Label lblTitle 
      Caption         =   "Total Cost (DC) && (IC)"
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
      Left            =   1200
      TabIndex        =   39
      Top             =   6600
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Caption         =   "Customer relation && administration"
      Height          =   255
      Index           =   14
      Left            =   4920
      TabIndex        =   38
      Top             =   6240
      Width           =   2535
   End
   Begin VB.Label lblTitle 
      Caption         =   "Downtime (production loss)"
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   37
      Top             =   5880
      Width           =   2265
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Transpotation costs"
      Height          =   195
      Index           =   12
      Left            =   4920
      TabIndex        =   36
      Top             =   5550
      Width           =   1875
   End
   Begin VB.Label lblTitle 
      Caption         =   "Training costs"
      Height          =   255
      Index           =   11
      Left            =   4920
      TabIndex        =   35
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Caption         =   "Rentals"
      Height          =   255
      Index           =   10
      Left            =   4920
      TabIndex        =   34
      Top             =   4800
      Width           =   2025
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-contracting costs"
      Height          =   195
      Index           =   9
      Left            =   4920
      TabIndex        =   33
      Top             =   4470
      Width           =   2235
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Clean-up costs"
      Height          =   195
      Index           =   8
      Left            =   4920
      TabIndex        =   32
      Top             =   3360
      Width           =   1875
   End
   Begin VB.Label lblTitle 
      Caption         =   "Disposal costs"
      Height          =   255
      Index           =   7
      Left            =   4920
      TabIndex        =   31
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Label lblTitle 
      Caption         =   "Overtime costs"
      Height          =   255
      Index           =   6
      Left            =   4920
      TabIndex        =   30
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label11 
      Caption         =   "Amounts (a) x (b) x (c)"
      Height          =   255
      Left            =   1440
      TabIndex        =   29
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "(c) Time spent on  investigation (hours)"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   5880
      Width           =   2775
   End
   Begin VB.Label Label9 
      Caption         =   "(b) Approximated hourly rate"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "(a) # of people involved in investigation"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Investigation costs"
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
      Left            =   240
      TabIndex        =   25
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "First Aid costs"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Material loss"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "b) Vehicle"
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Damage costs    a) Property"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "INDIRECT COSTS (IC)"
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
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Total Cost (DC)"
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Caption         =   "Sickness and accident costs"
      Height          =   255
      Index           =   3
      Left            =   5280
      TabIndex        =   18
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Workers compendation (actual costs)"
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
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   3315
   End
   Begin VB.Label lblTitle 
      Caption         =   "DIRECT COSTS (DC)"
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
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   2145
   End
   Begin VB.Label lblTitle 
      Caption         =   "Unemployment and disability costs"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   2865
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Workers compensation (actual costs)"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   2430
      Width           =   2835
   End
   Begin VB.Label lblTitle 
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
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label lblEEID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "AC_Empnbr"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7350
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "AC_CompNo"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6210
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEHSCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewCode
Dim blError As Boolean
Dim rsDATA As New ADODB.Recordset
Dim fGLBNew As Boolean

Function chkHSCost()

Dim SQLQ As String, Msg As String, dd#

chkHSCost = False

On Error GoTo chkHSCost_Err

Dim tTime As Variant
Dim Part1$, Part2$

'~~

If Len(txtShift) < 1 Then
    MsgBox "Incident Number is a required field"
    comShift.SetFocus
    Exit Function
End If
If Not IfIncidentNo(Val(txtShift)) Then
    MsgBox "Incident Number Not Valid"
    comShift.SetFocus
    Exit Function
End If


chkHSCost = True

Exit Function

chkHSCost_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "WFC_Accident_Cost", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function
Sub Tot_DI()
    medTotDC = Val(medCompe) + Val(medUnemp) + Val(medSick)
    Call Tot '#3975
End Sub
Sub Tot_IC()
    medTotIC = Val(medPrope) + Val(medVehic) + Val(medMater) + Val(medFirst) + Val(medAmount)
    medTotIC = Val(medTotIC) + Val(medClean) + Val(medDispo) + Val(medOTime) + Val(medSubco) + Val(medRent)
    medTotIC = Val(medTotIC) + Val(medTrain) + Val(medTrans) + Val(medDown) + Val(medCust)
    Call Tot '#3975
End Sub
Sub Tot()
    medTotal = Val(medTotDC) + Val(medTotIC)
End Sub
Sub Tot_Amount()
    medAmount = Val(medPeopl) * Val(medHourly) * Val(medTime)
    Call Tot_IC
End Sub
Public Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err
fGLBNew = False
Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call SET_UP_MODE
Me.vbxTrueGrid.Enabled = True
'Call ST_UPD_MODE(False)  ' reset screen's attributes
'If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
'Me.vbxTrueGrid.SetFocus
'End If
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "WFC_Accident_Cost", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub


Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEHSCOST" Then glbOnTop = ""

End Sub

Private Sub cmdContact_Click()
frmEHSContact.Show
Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, X

If Not gSec_Upd_HSCost Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call Display_Value
Me.vbxTrueGrid.SetFocus
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "WFC_Accident_Cost", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub


Private Sub cmdIncident_Click()
frmEHSINCIDENT.Show
Unload Me
End Sub

Sub cmdInjLoc_Click()
frmEHSINJURY.Show
Unload Me
End Sub

Public Sub cmdModify_Click()

If Not gSec_Upd_HSCost Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

On Error GoTo Mod_Err
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "WFC_Accident_Cost", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_HSCost Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If
fGLBNew = True
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

On Error GoTo AddN_Err
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    'Me.vbxTrueGrid.Enabled = False
End If
Me.vbxTrueGrid.Enabled = False
Data1.Recordset.AddNew

'If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblEEID = glbLEE_ID
lblCNum.Caption = "001"
'#3975
medAmount = 0
medTotal = 0
medTotDC = 0
medTotIC = 0

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "WFC_Accident_Cost", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Public Sub cmdOK_Click()
Dim X
On Error GoTo Add_Err

If Not chkHSCost() Then Exit Sub

If glbtermopen Then Data1.Recordset("TERM_SEQ") = glbTERM_Seq
Call UpdUStats(Me) ' update user's stats (who did it and when)
'#3975
'Data1.Recordset("AC_Case") = txtShift & ""
Data1.Recordset("AC_Empnbr") = lblEEID
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call Display_Value
fGLBNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
Me.vbxTrueGrid.Enabled = True

'Me.vbxTrueGrid.SetFocus
If NextFormIF("Corrective Action") Then
    Call cmdNew_Click
End If
Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "WFC_Accident_Cost", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Corrective Actions"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Corrective Actions"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdTCause_Click()
frmEHSCause.Show
Unload Me
End Sub


Sub cmdWCBMed_Click()
frmEHSWCB.Show
Unload Me
End Sub

Private Sub cmdWSIB_Click()
frmEHSWCBC.Show
Unload Me
End Sub

Sub comShift_Change()
'txtShift = comShift  'JDY
End Sub

Sub comShift_Click()
'txtShift = comShift      'JDY
End Sub


Private Sub comShift_LostFocus()
txtShift = comShift  'JDY
End Sub

'Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = ErrorNumber
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "WFC_Accident_Cost", "SELECT")
'
'
'End Sub




Function EERetrieve()


Dim SQLQ As String

EERetrieve = False
blError = False
Screen.MousePointer = HOURGLASS
On Error GoTo EERError


'If glbtermopen Then         'Lucy July 5, 2000
'    SQLQ = "Select * from Term_OHS_CORRECTIVE"
'    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'    SQLQ = SQLQ & " ORDER BY CR_CASE DESC"
'Else
    SQLQ = "Select * from WFC_Accident_Cost"
    SQLQ = SQLQ & " where AC_Empnbr = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY AC_Case DESC"
'End If

Data1.RecordSource = SQLQ
Data1.Refresh

'If glbtermopen Then     'Lucy July 5, 2000
'    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
'    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
'    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
'Else
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
'End If

Data3.RecordSource = SQLQ
Data3.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
blError = True
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "WFC_Accident_Cost", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function





End Function





Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMEHSCOST"
Call EERetrieve
End Sub

Sub Form_GotFocus()
glbOnTop = "FRMEHSCOST"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ1
glbOnTop = "FRMEHSCOST"


Data1.ConnectionString = glbAdoIHRDB 'glbAdoIHRWFC
Data3.ConnectionString = glbAdoIHRDB

If glbLEE_ID = 0 Then frmEEFIND.Show 1
If glbLEE_ID = 0 Then Unload Me: Exit Sub

If EERetrieve() = False Then
    If blError Then
        Unload Me
        Exit Sub
    Else
        MsgBox "Sorry, Employee can not be found"
        If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
    End If
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS

comShift.Clear

Do Until Data3.Recordset.EOF                  'JDY
  comShift.AddItem Data3.Recordset("EC_CASE") 'JDY
  Data3.Recordset.MoveNext                    'JDY
Loop

Me.vbxTrueGrid.SetFocus
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Accident Cost Analysis Data - " & Left$(glbLEE_SName, 8)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value

fGLBNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

If Not gSec_Upd_HSCost Then
    'cmdModify.Enabled = False
    'cmdNew.Enabled = False
    'cmdDelete.Enabled = False
End If
Screen.MousePointer = DEFAULT

End Sub

Sub Form_LostFocus()
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

Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSCost = Nothing 'carmen 18 may 00
Call NextForm
End Sub
Function IfIncidentNo(InciNo As Double)
  IfIncidentNo = False
  If Data3.Recordset.BOF And Data3.Recordset.EOF Then
     Exit Function
  End If
  Data3.Recordset.MoveFirst
  Data3.Recordset.Find "EC_Case=" & InciNo
  If Data3.Recordset.EOF Then Exit Function
  IfIncidentNo = True
End Function


Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

comShift.Enabled = TF
medCompe.Enabled = TF
medUnemp.Enabled = TF
medPrope.Enabled = TF
medVehic.Enabled = TF
medMater.Enabled = TF
medFirst.Enabled = TF
medPeopl.Enabled = TF
medHourly.Enabled = TF
medTime.Enabled = TF
medSick.Enabled = TF
medClean.Enabled = TF
medDispo.Enabled = TF
medOTime.Enabled = TF
medSubco.Enabled = TF
medRent.Enabled = TF
medTrain.Enabled = TF
medTrans.Enabled = TF
medDown.Enabled = TF
medCust.Enabled = TF

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'cmdWCBMed.Enabled = FT
'cmdIncident.Enabled = FT
'cmdTCause.Enabled = FT
'cmdContact.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdWSIB.Enabled = FT
'vbxTrueGrid.Enabled = FT
'If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
'Else
'    cmdModify.Enabled = True
'End If
End Sub

Private Sub medAmount_LostFocus()
    If Not IsNumeric(medAmount) Then medAmount = 0
    Call Tot_IC
End Sub

Private Sub medClean_LostFocus()
    If Not IsNumeric(medClean) Then medClean = 0
    Call Tot_IC
End Sub

Private Sub medClean_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCompe_LostFocus()
    If Not IsNumeric(medCompe) Then medCompe = 0
    Call Tot_DI
End Sub

Private Sub medCompe_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCust_LostFocus()
    If Not IsNumeric(medCust) Then medCust = 0
    Call Tot_IC
End Sub

Private Sub medCust_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medDispo_LostFocus()
    If Not IsNumeric(medDispo) Then medDispo = 0
    Call Tot_IC
End Sub

Private Sub medDispo_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medDown_LostFocus()
    If Not IsNumeric(medDown) Then medDown = 0
    Call Tot_IC
End Sub

Private Sub medDown_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medFirst_LostFocus()
    If Not IsNumeric(medFirst) Then medFirst = 0
    Call Tot_IC
End Sub

Private Sub medFirst_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medHourly_LostFocus()
    If Not IsNumeric(medHourly) Then medHourly = 0
    Call Tot_Amount
End Sub

Private Sub medHourly_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medMater_LostFocus()
    If Not IsNumeric(medMater) Then medMater = 0
    Call Tot_IC
End Sub

Private Sub medMater_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medOTime_LostFocus()
    If Not IsNumeric(medOTime) Then medOTime = 0
    Call Tot_IC
End Sub

Private Sub medOTime_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPeopl_LostFocus()
    If Not IsNumeric(medPeopl) Then medPeopl = 0
    'medPeopl = Round(Val(medPeopl), 0)
    Call Tot_Amount
End Sub

Private Sub medPeopl_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPrope_LostFocus()
    If Not IsNumeric(medPrope) Then medPrope = 0
    Call Tot_IC
End Sub

Private Sub medPrope_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medRent_LostFocus()
    If Not IsNumeric(medRent) Then medRent = 0
    Call Tot_IC
End Sub

Private Sub medRent_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSick_LostFocus()
    If Not IsNumeric(medSick) Then medSick = 0
    Call Tot_DI
End Sub

Private Sub medSick_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSubco_LostFocus()
    If Not IsNumeric(medSubco) Then medSubco = 0
    Call Tot_IC
End Sub

Private Sub medSubco_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTime_LostFocus()
    If Not IsNumeric(medTime) Then medTime = 0
    Call Tot_Amount
End Sub

Private Sub medTime_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTotDC_LostFocus()
    If Not IsNumeric(medTotDC) Then medTotDC = 0
    Call Tot
End Sub

Private Sub medTotIC_LostFocus()
    If Not IsNumeric(medTotIC) Then medTotIC = 0
    Call Tot
End Sub

Private Sub medTrain_LostFocus()
    If Not IsNumeric(medTrain) Then medTrain = 0
    Call Tot_IC
End Sub

Private Sub medTrain_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTrans_LostFocus()
    If Not IsNumeric(medTrans) Then medTrans = 0
    Call Tot_IC
End Sub

Private Sub medTrans_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medUnemp_LostFocus()
    If Not IsNumeric(medUnemp) Then medUnemp = 0  '--
    Call Tot_DI
End Sub

Private Sub medUnemp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medVehic_LostFocus()
    If Not IsNumeric(medVehic) Then medVehic = 0
    Call Tot_IC
End Sub

Private Sub medVehic_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtShift_Change()
 
  If Not (Val(txtShift) = 0) Then
    comShift = txtShift
  Else
    comShift = ""
  End If

End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        'End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
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
        
        SQLQ = "Select * from WFC_Accident_Cost"
        SQLQ = SQLQ & " where AC_Empnbr = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    'If cmdOK.Enabled Then
    '    cmdOK.SetFocus
    'Else
    '    cmdModify.SetFocus
    'End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'#3975
Call Tot_DI
Call Tot_Amount
Call Tot_IC
Call Tot
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_HSCost
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

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fGLBNew Then
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
Call ST_UPD_MODE(TF)
End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        'Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        'rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsDATA.Open Data1.RecordSource, gdbAdoIhrWFC, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM WFC_Salary_Administration " & " ORDER BY MarketLine, [Band]"
        
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    'rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsDATA.Open SQLQ, gdbAdoIhrWFC, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    'Call Set_Control("R", Me, rsDATA)

End Sub
