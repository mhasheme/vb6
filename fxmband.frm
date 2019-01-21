VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMBand 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary Grids"
   ClientHeight    =   8475
   ClientLeft      =   105
   ClientTop       =   645
   ClientWidth     =   11535
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8475
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmband.frx":0000
      Height          =   2595
      Left            =   0
      OleObjectBlob   =   "fxmband.frx":0014
      TabIndex        =   22
      Top             =   120
      Width           =   10515
   End
   Begin VB.PictureBox frmDetails 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4530
      Left            =   0
      ScaleHeight     =   4530
      ScaleWidth      =   10185
      TabIndex        =   0
      Top             =   2280
      Width           =   10185
      Begin MSMask.MaskEdBox MskDollars 
         DataField       =   "LDollars"
         DataSource      =   "data1"
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   3
         Tag             =   "01-Low Dollars"
         Top             =   1290
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDollars 
         DataField       =   "MDollars"
         DataSource      =   "data1"
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   4
         Tag             =   "01-MidPoint Dollars"
         Top             =   1650
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskDollars 
         DataField       =   "HDollars"
         DataSource      =   "data1"
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   5
         Tag             =   "01-High Dollars"
         Top             =   2010
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFiscalYear 
         DataField       =   "FiscalYear"
         DataSource      =   "data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Tag             =   "01-High Dollars"
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "SectionCode"
         Height          =   285
         Index           =   0
         Left            =   1845
         TabIndex        =   12
         Tag             =   "00-Section - Code"
         Top             =   3840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "BAND"
         Height          =   285
         Index           =   2
         Left            =   1845
         TabIndex        =   1
         Tag             =   "00-Band - Code"
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFBD"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "MarketLine"
         Height          =   285
         Index           =   3
         Left            =   1845
         TabIndex        =   2
         Tag             =   "00-Market Line - Code"
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFML"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "CurrencyIndicator"
         Height          =   285
         Index           =   4
         Left            =   1845
         TabIndex        =   7
         Tag             =   "00-Currency Indicator - Code"
         Top             =   2760
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFCI"
      End
      Begin MSMask.MaskEdBox MskMidPointPer 
         DataField       =   "MIDPOINT_PER"
         DataSource      =   "data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Tag             =   "10-Percentage of MidPoint"
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   5
         Left            =   6960
         TabIndex        =   14
         Tag             =   "00-Union"
         Top             =   3840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
         MaxLength       =   15
      End
      Begin MSMask.MaskEdBox MskFiscalYe2 
         DataSource      =   "data1"
         Height          =   315
         Left            =   7275
         TabIndex        =   13
         Tag             =   "01-High Dollars"
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskTROIC 
         DataSource      =   "data1"
         Height          =   315
         Left            =   9960
         TabIndex        =   10
         Tag             =   "01-High Dollars"
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskACorpObj 
         DataSource      =   "data1"
         Height          =   315
         Left            =   9960
         TabIndex        =   9
         Tag             =   "01-High Dollars"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
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
         Format          =   "0.00%"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskApprLimit 
         DataField       =   "APPR_LIMIT"
         DataSource      =   "data1"
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Tag             =   "01-High Dollars"
         Top             =   3100
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Format          =   "#,##0;(#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label lblApprLimit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Signing Approval Limit"
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
         Left            =   240
         TabIndex        =   34
         Top             =   3100
         Width           =   1695
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ROIC"
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
         Left            =   8280
         TabIndex        =   33
         Top             =   2520
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblIPPerc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Incentive Percentage"
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
         Left            =   8280
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year Filter"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5565
         TabIndex        =   31
         Top             =   3480
         Width           =   885
      End
      Begin VB.Label lblUnionFilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant  Filter"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5565
         TabIndex        =   30
         Top             =   3840
         Width           =   990
      End
      Begin VB.Label lblMidPPer 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "% of Salary"
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
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label lblPlant 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   3840
         Width           =   1260
      End
      Begin VB.Label lblFiscalYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3480
         Width           =   960
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "MidPoint Jobrate"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1650
         Width           =   1440
      End
      Begin VB.Label lblCurrencyIndicator 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Indicator"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1290
      End
      Begin VB.Label lblMarketLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   930
         Width           =   1020
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label lblBand 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Band"
         DataSource      =   "Data1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   570
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdDelDupRec 
      Caption         =   "Delete Record"
      Height          =   375
      Left            =   7920
      TabIndex        =   29
      Top             =   7200
      Width           =   2055
   End
   Begin VB.CommandButton cmdCopyTo 
      Caption         =   "Copy To"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   7200
      Width           =   1335
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   21
      Top             =   7815
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6465
         Top             =   180
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   450
      Left            =   9720
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   794
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   3645
      TabIndex        =   26
      Tag             =   "00-Section - Code"
      Top             =   7250
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblDPlant 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Plant "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2040
      TabIndex        =   27
      Top             =   7290
      Width           =   1620
   End
End
Attribute VB_Name = "frmMBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'True DBGrid changed
Option Explicit
Dim fglbRecords%, fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer, fglbNew
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim xID
Dim OldLimit
Dim xLastFisYear

Private Function chkBand()
Dim SQLQ As String, Msg As String, dd#, PID&, Factor$
chkBand = False
On Error GoTo chkPosEval_Err
If Trim(clpCode(2)) = "" Then
    MsgBox "Band is a required field"
    'cmbBand.SetFocus
    Exit Function
End If
If Trim(clpCode(3)) = "" Then
    MsgBox "Market Line is a required field"
    'cmbMarketLine.SetFocus
    Exit Function
End If

If Val(MskDollars(0)) = 0 Then
    MsgBox "Low Dollars must be greater than 0 "
    MskDollars(0).SetFocus
    Exit Function
End If
If Val(MskDollars(1)) = 0 Then
    MsgBox "MidPoint Dollars must be greater than 0 "
    MskDollars(1).SetFocus
    Exit Function
End If
If Val(MskDollars(2)) = 0 Then
    MsgBox "High Dollars must be greater than 0 "
    MskDollars(2).SetFocus
    Exit Function
End If

If Len(MskFiscalYear.Text) > 0 Then
    If Not IsNumeric(MskFiscalYear.Text) Then
        MsgBox "Invalid Fiscal Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYear.SetFocus
    Exit Function
End If
If Len(clpCode(0).Text) > 0 Then
    If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Plant is a required field"
    clpCode(0).SetFocus
    Exit Function
End If

If modISDupFactor(glbPos$, Factor$, PID&) And fglbNew Then
    MsgBox "[Band + Market Line + Fiscal Year + Plant] must be unique"
    'cmbBand.SetFocus
    Exit Function
End If
If Val(MskDollars(1)) < Val(MskDollars(0)) Then
    MsgBox "MidPoint Dollars must be greater than Low Dollars"
    MskDollars(1).SetFocus
    Exit Function
End If

If Val(MskDollars(2)) < Val(MskDollars(1)) Then
    MsgBox "High Dollars must be greater than MidPoint Dollars"
    MskDollars(2).SetFocus
    Exit Function
End If

If Len(MskApprLimit.Text) > 0 Then
    If Not IsNumeric(MskApprLimit.Text) Then
            MsgBox "Invalid Signing Approval Limit."
            MskApprLimit.SetFocus
        Exit Function
    End If
End If
chkBand = True

Exit Function

chkPosEval_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBEVL", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub clpCode_Change(Index As Integer)
    If Index = 3 Then
        clpCode(4) = Left(clpCode(3).Text, 2)
    End If

    If Index = 5 Then
        If Not clpCode(5).Caption = "Unassigned" Then
            Call EERetrieve
        End If
    End If
End Sub

Private Sub clpCode_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub cmbBand_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbBand_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


'Private Sub cmbMarketLine_Change()
''clpCode(3) = Left(cmbMarketLine.Text, 4)
''txtCurrencyIndicator = Left(cmbMarketLine, 2)
'End Sub

'Private Sub cmbMarketLine_click()
''clpCode(3) = Left(cmbMarketLine.Text, 4)
''txtCurrencyIndicator = Left(cmbMarketLine, 2)
'End Sub

'Private Sub cmbMarketLine_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmbMarketLine_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

Public Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err
fglbNew = False
Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Data1.Refresh
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)  ' reset screen's attributes
'cmbBand.Enabled = False
'cmbMarketLine.Enabled = False
clpCode(2).Enabled = False
clpCode(3).Enabled = False

Me.vbxTrueGrid.SetFocus
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub


Public Sub cmdClose_Click()
Unload Me
End Sub



Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, SQLQ, X%, xEmpnbr
Dim xBand, xMarketline, DeleteRight, xJob
'Dim XTB As Recordset

Dim snapAssBand As New ADODB.Recordset
If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

Screen.MousePointer = HOURGLASS
SQLQ = "SELECT SH_EMPNBR,ED_SURNAME,ED_FNAME FROM HR_SALARY_HISTORY,HREMP "
SQLQ = SQLQ & " WHERE SH_BAND = '" & Data1.Recordset("Band") & "'"
SQLQ = SQLQ & " AND SH_MARKETLINE = '" & Data1.Recordset("MarketLine") & "'"
SQLQ = SQLQ & " AND ED_EMPNBR=SH_EMPNBR order by SH_EMPNBR"

If snapAssBand.State <> 0 Then snapAssBand.Close
snapAssBand.Open SQLQ, gdbAdoIhr001, adOpenKeyset
Screen.MousePointer = DEFAULT

If Not (snapAssBand.BOF And snapAssBand.EOF) Then
    X% = 0: xEmpnbr = 0
    Msg = "This record is in the following employees'" & Chr(10) & "salary history:"
    While Not snapAssBand.EOF And X% < 10
        If xEmpnbr <> snapAssBand("sh_EMPNBR") Then
          Msg = Msg & Chr(10) & snapAssBand("ED_surname") & ", " & snapAssBand("ED_FName") & " -  # " & snapAssBand("sh_EMPNBR")
          X% = X% + 1
        End If
        xEmpnbr = snapAssBand("sh_EMPNBR")
        snapAssBand.MoveNext
    Wend
    Msg = Msg & Chr(10) & "Record will not be deleted."
    MsgBox Msg
    Exit Sub
End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
SQLQ = "Delete FROM WFC_Salary_Administration "
SQLQ = SQLQ & "where [band]='" & Trim(clpCode(2)) & "'"
SQLQ = SQLQ & "and MarketLine='" & Trim(clpCode(3)) & "'"
Data1.Recordset.ActiveConnection.Execute SQLQ
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call Display_Value

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBEVL", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub


Public Sub cmdModify_Click()

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

fglbEditMode% = True

On Error GoTo Mod_Err
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

'cmbBand.Enabled = False
'cmbMarketLine.Enabled = False
clpCode(2).Enabled = False
clpCode(3).Enabled = False
'MskDollars(0).SetFocus
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
Dim SQLQ As String
If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If
'Call SET_UP_MODE
'Call ST_UPD_MODE(True)

On Error GoTo AddN_Err

Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
'Call Set_Control("B", Me)
'rsDATA.AddNew

fglbEditMode% = True
fglbNew = True
Call SET_UP_MODE
'cmbBand.Enabled = True
'cmbMarketLine.Enabled = True
clpCode(2).Enabled = True
clpCode(3).Enabled = True

'cmbBand.SetFocus
'cmbBand = ""
'cmbMarketLine = ""
clpCode(2).Text = ""
clpCode(3).Text = ""
'txtBand.Enabled = False
'clpCode(3).Enabled = False

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRJOBEVL", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub


Public Sub cmdOK_Click()
Dim a As Integer, Msg$, INo&, X%
Dim xTot As Integer

On Error GoTo OK_Err

If Not chkBand() Then Exit Sub
'cmbCurrencyIndicator_setup2 Me
'setMarketLine Me
If Len(clpCode(0).Text) > 0 Then Data1.Recordset("SectionCode") = clpCode(0).Text
Data1.Recordset("BAND") = clpCode(2) 'txtBand
If Len(clpCode(3).Text) > 0 Then Data1.Recordset("MarketLine") = clpCode(3)
If Len(clpCode(4).Text) > 0 Then Data1.Recordset("CurrencyIndicator") = clpCode(4)
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
xID = Data1.Recordset("ID")
Data1.Refresh
Data1.Recordset.Find "ID= " & xID

fglbNew = False
fglbEditMode% = False
Call Display_Value
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
'cmbBand.Enabled = False
'cmbMarketLine.Enabled = False
clpCode(2).Enabled = False
clpCode(3).Enabled = False
Me.vbxTrueGrid.SetFocus

If MskFiscalYear.Text = xLastFisYear Then 'Ticket #29846 Franks 03/06/2017
    If Not (OldLimit = MskApprLimit.Text) Then
        If IsNumeric(MskApprLimit.Text) Then
            Msg$ = "Signing Approval Limit has been changed. "
            Msg$ = Msg$ & Chr(10) & "Are you sure you want to update the Position Master with the new Signing Approval Limit? "
            a% = MsgBox(Msg$, 36, "Confirm Update")
            If a% <> 6 Then
            Else 'Yes
                xTot = WFCSigningApprovalLimitUpt(clpCode(0).Text, clpCode(2).Text, clpCode(3).Text, MskApprLimit.Text) 'Plant, Band, Marketline, Limit
                If xTot = 0 Then
                    MsgBox "No matching Position Master record found."
                Else
                    MsgBox xTot & " Position Master record(s) updated."
                End If
            End If
        End If
    End If
End If

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBEVL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Unload Me


End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub


Private Sub clpCode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 5 Then
        If Not clpCode(5).Caption = "Unassigned" Then
            Call EERetrieve
        End If
    End If
End Sub

Private Sub cmdCopyTo_Click()
Dim SQLQ
Dim rsFBand As New ADODB.Recordset
Dim a As Integer, Msg$, INo&, X%

    If Len(MskFiscalYear.Text) > 0 Then
        If Not IsNumeric(MskFiscalYear.Text) Then
            MsgBox "Invalid Fiscal Year."
            MskFiscalYear.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Fiscal Year is a required field"
        MskFiscalYear.SetFocus
        Exit Sub
    End If
    If Len(clpCode(0).Text) > 0 Then
        If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
            MsgBox "If code entered it must be known"
            clpCode(0).SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Plant is a required field"
        clpCode(0).SetFocus
        Exit Sub
    End If

    If Len(clpCode(1).Text) > 0 Then
        If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
            MsgBox "If code entered it must be known"
            clpCode(1).SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Destination Plant is a required field"
        clpCode(1).SetFocus
        Exit Sub
    End If
    
    If clpCode(0).Text = clpCode(1).Text Then
        MsgBox "Destination Plant Code is equal to From Plant Code"
        clpCode(1).SetFocus
        Exit Sub
    End If

    SQLQ = "SELECT * FROM WFC_Salary_Administration WHERE BAND = '" & clpCode(2) & "' "
    SQLQ = SQLQ & "AND MarketLine = '" & clpCode(3) & "' "
    SQLQ = SQLQ & "AND FiscalYear = " & MskFiscalYear & " "
    SQLQ = SQLQ & "AND SectionCode = '" & clpCode(1) & "' "
    rsFBand.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsFBand.EOF Then
        rsFBand.Close
        MsgBox "Duplicate record found."
        clpCode(1).SetFocus
        Exit Sub
    End If



    Msg$ = "Are You Sure You Want To Copy this record? "
    a% = MsgBox(Msg$, 36, "Confirm Copy")
    If a% <> 6 Then
        rsFBand.Close
        Exit Sub
    End If
    rsFBand.AddNew
    rsFBand("BAND") = clpCode(2)
    rsFBand("MarketLine") = clpCode(3)
    If Len(MskDollars(0).Text) > 0 Then rsFBand("LDollars") = MskDollars(0).Text
    If Len(MskDollars(1).Text) > 0 Then rsFBand("MDollars") = MskDollars(1).Text
    If Len(MskDollars(2).Text) > 0 Then rsFBand("HDollars") = MskDollars(2).Text
    If Len(MskMidPointPer.Text) > 0 Then rsFBand("MIDPOINT_PER") = MskMidPointPer.Text
    If Len(clpCode(4).Text) > 0 Then rsFBand("CurrencyIndicator") = clpCode(4).Text
    rsFBand("FiscalYear") = MskFiscalYear
    rsFBand("SectionCode") = clpCode(1)
    'If Len(MskACorpObj.Text) > 0 Then rsFBand("IP_PERCENTAGE") = MskACorpObj.Text 'Ticket #29014 Franks 09/06/2016
    'If Len(MskTROIC.Text) > 0 Then rsFBand("IP_ROIC") = MskTROIC.Text 'Ticket #29014 Franks 09/06/2016
    If Len(MskApprLimit.Text) > 0 Then rsFBand("APPR_LIMIT") = MskApprLimit.Text
    rsFBand.Update
    xID = rsFBand("ID")
    rsFBand.Close
    Data1.Refresh
    Data1.Recordset.Find "ID= " & xID
    Call Display_Value

End Sub

Private Sub cmdDelDupRec_Click()
Dim a As Integer, Msg As String, SQLQ, X%, xEmpnbr
Dim xBand, xMarketline, DeleteRight, xJob
'Dim XTB As Recordset

Dim snapAssBand As New ADODB.Recordset
If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

Screen.MousePointer = HOURGLASS

Msg = "This function is only for deleting duplicate record"
Msg = Msg & Chr(10) & "Are You Sure You Want To Delete This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If
xID = Data1.Recordset("ID")
SQLQ = "Delete FROM WFC_Salary_Administration "
SQLQ = SQLQ & "where [ID]=" & xID & " "
Data1.Recordset.ActiveConnection.Execute SQLQ
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call Display_Value

Call SET_UP_MODE
'Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBEVL", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRJOBEVL", "SELECT")


End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
glbOnTop = "FRMMBAND"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS
Me.Caption = "Salary Grids"

Data1.ConnectionString = glbAdoIHRDB 'glbAdoIHRWFC
'Data1.RecordSource = "WFC_Salary_Administration"



Screen.MousePointer = DEFAULT
X% = EERetrieve()

'Band_AddItem Me
'MarketLine_AddItem Me
'CurrencyIndicator_AddItem Me
fglbNew = False

Call Display_Value

Screen.MousePointer = HOURGLASS
Call INI_Controls(Me)
'Me.vbxTrueGrid.SetFocus

xLastFisYear = getLastFisYear 'Ticket #29846 Franks 03/06/2017

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


End Sub

Private Function EERetrieve()
Dim SQLQ
EERetrieve = False
Screen.MousePointer = HOURGLASS
On Error GoTo modGetPosEvalsErr

SQLQ = "SELECT * FROM WFC_Salary_Administration "
SQLQ = SQLQ & "WHERE (1=1) "
If Len(glbPlantCode) > 0 Then
    SQLQ = SQLQ & "AND SectionCode = '" & glbPlantCode & "' "
End If
If Len(clpCode(5).Text) > 0 Then
    SQLQ = SQLQ & "AND SectionCode = '" & clpCode(5).Text & "' "
End If
If Len(MskFiscalYe2.Text) > 0 Then
    If IsNumeric(MskFiscalYe2.Text) Then
        SQLQ = SQLQ & "AND FiscalYear = " & MskFiscalYe2.Text & " "
    End If
End If
SQLQ = SQLQ & "ORDER BY SectionCode,FiscalYear,MarketLine, [Band] "
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function



modGetPosEvalsErr:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Pos Skills", "HRJOBSK", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function





End Function

Private Function modISDupFactor(Pos$, Factor$, PID&)
Dim SQLQ As String
Dim snapEval As New ADODB.Recordset

modISDupFactor = True

On Error GoTo modISDupFactor_Err
Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM WFC_Salary_Administration "
SQLQ = SQLQ & "where [band]='" & Trim(clpCode(2)) & "'"
SQLQ = SQLQ & " and MarketLine='" & Trim(clpCode(3)) & "'"
SQLQ = SQLQ & " and SectionCode='" & Trim(clpCode(0)) & "'"
SQLQ = SQLQ & " and FiscalYear='" & MskFiscalYear & "'"

snapEval.Open SQLQ, gdbAdoIhrWFC

If snapEval.BOF And snapEval.EOF Then
    modISDupFactor = False
End If

snapEval.Close
Screen.MousePointer = DEFAULT

Exit Function

modISDupFactor_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Function

Private Sub MskACorpObj_GotFocus()
If IsNumeric(MskACorpObj.Text) Then
    MskACorpObj.Text = MskACorpObj.Text * 100
End If
End Sub

Private Sub MskACorpObj_LostFocus()
If IsNumeric(MskACorpObj.Text) Then
    MskACorpObj.Text = MskACorpObj.Text / 100
End If
End Sub

Private Sub MskDollars_GotFocus(Index As Integer)
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

glbOHSEdit% = TF

fUPMode = TF    ' update mode
frmDetails.Enabled = TF


'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'vbxTrueGrid.Enabled = FT
'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
'    cmdDelete.Enabled = False
'    cmdModify.Enabled = False
'End If

End Sub





Private Sub MskFiscalYe2_Change()
    If Not clpCode(5).Caption = "Unassigned" Then
        Call EERetrieve
    End If
End Sub

Private Sub MskFiscalYe2_KeyUp(KeyCode As Integer, Shift As Integer)
        If Not clpCode(5).Caption = "Unassigned" Then
            Call EERetrieve
        End If
End Sub

Private Sub MskMidPointPer_GotFocus()
If IsNumeric(MskMidPointPer) Then
    MskMidPointPer = MskMidPointPer * 100
End If
End Sub

Private Sub MskMidPointPer_LostFocus()
If IsNumeric(MskMidPointPer) Then
    MskMidPointPer = MskMidPointPer / 100
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
        
        SQLQ = "SELECT * FROM WFC_Salary_Administration "
        'If Len(glbPlantCode) > 0 Then
        '    SQLQ = SQLQ & "WHERE SectionCode = '" & glbPlantCode & "' "
        'End If
        SQLQ = SQLQ & "WHERE (1=1) "
        If Len(glbPlantCode) > 0 Then
            SQLQ = SQLQ & "AND SectionCode = '" & glbPlantCode & "' "
        End If
        If Len(clpCode(5).Text) > 0 Then
            SQLQ = SQLQ & "AND SectionCode = '" & clpCode(5).Text & "' "
        End If
        If Len(MskFiscalYe2.Text) > 0 Then
            SQLQ = SQLQ & "AND FiscalYear = " & MskFiscalYe2.Text & " "
        End If
        
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    'If cmdOK.Enabled Then
    '    cmdOK.SetFocus
    'Else
    '    cmdClose.SetFocus
    'End If
End If

End Sub



Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  Call Display_Value
  OldLimit = MskApprLimit.Text
  
  'cmbBand_SETUP Me
  'cmbCurrencyIndicator_setup2 Me
  'setMarketLine Me
  'MarketLine_Desc Me
  
  
End Sub
''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        'rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsDATA.Open Data1.RecordSource, gdbAdoIhrWFC, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM WFC_Salary_Administration "
    If Data1.Recordset("ID") > 0 Then
        SQLQ = SQLQ & "WHERE ID = " & Data1.Recordset("ID") & " "
    End If
    SQLQ = SQLQ & " ORDER BY MarketLine, [Band] "
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    'rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsDATA.Open SQLQ, gdbAdoIhrWFC, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    'Call Set_Control("R", Me, rsDATA)
    If Not IsNull(rsDATA("SectionCode")) Then
        clpCode(0).Text = rsDATA("SectionCode")
    Else
        clpCode(0).Text = ""
    End If
    If Not IsNull(rsDATA("BAND")) And Not fglbNew Then
        clpCode(2).Text = rsDATA("BAND")
    Else
        clpCode(2).Text = ""
    End If
    If Not IsNull(rsDATA("MarketLine")) And Not fglbNew Then
        clpCode(3).Text = rsDATA("MarketLine")
    Else
        clpCode(3).Text = ""
    End If
    If Not IsNull(rsDATA("CurrencyIndicator")) And Not fglbNew Then
        clpCode(4).Text = rsDATA("CurrencyIndicator")
    Else
        clpCode(4).Text = ""
    End If
    
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property
Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelatePOS
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Inq_SalaryGrids
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
Call ST_UPD_MODE(TF)
End Sub

Private Function getLastFisYear() 'Ticket #29846 Franks 03/06/2017
Dim rsTmp As New ADODB.Recordset
Dim SQLQ
Dim retval
    retval = ""
    SQLQ = "SELECT TOP 1 * FROM WFC_Salary_Administration ORDER BY FiscalYear DESC "
    rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmp.EOF Then
        If Not IsNull(rsTmp("FiscalYear")) Then
            retval = rsTmp("FiscalYear")
        End If
    End If
    rsTmp.Close
    getLastFisYear = retval
End Function
