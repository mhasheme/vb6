VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmECounsel 
   Caption         =   "Counselling"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   11055
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStreamRules 
      Caption         =   "Rules..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   2280
      TabIndex        =   41
      Tag             =   "Select File to Import"
      Top             =   6472
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cmbStream 
      Height          =   315
      ItemData        =   "fECounsel.frx":0000
      Left            =   3600
      List            =   "fECounsel.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "40-Select Stream"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtStream 
      Appearance      =   0  'Flat
      DataField       =   "CL_STREAM"
      Height          =   285
      Left            =   1800
      TabIndex        =   39
      ToolTipText     =   "Enter the Rule # violated"
      Top             =   6480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox memEmpResponse 
      Appearance      =   0  'Flat
      DataField       =   "CL_EMP_RESPONSE"
      Height          =   495
      Left            =   120
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "00-Employee Response"
      Top             =   5400
      Width           =   8775
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   7980
      TabIndex        =   30
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin INFOHR_Controls.EmployeeLookup elpCouByShow 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "10-Employee Number of individual's supervisor"
      Top             =   3600
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fECounsel.frx":002C
      Height          =   1875
      Left            =   0
      OleObjectBlob   =   "fECounsel.frx":0040
      TabIndex        =   0
      Top             =   600
      Width           =   8895
   End
   Begin INFOHR_Controls.DateLookup dlpCouDate 
      DataField       =   "CL_COUDATE"
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "41-Counselling Date"
      Top             =   2940
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CL_REASON"
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Tag             =   "01-Counselling Reason- Code"
      Top             =   3270
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CERE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CL_TYPE"
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Tag             =   "01-Counselling Type- Code"
      Top             =   2610
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CETY"
   End
   Begin VB.TextBox txtCouBy 
      Appearance      =   0  'Flat
      DataField       =   "CL_COUBY"
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "CL_COMMENTS"
      Height          =   855
      Left            =   120
      MaxLength       =   4000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "00-Comments"
      Top             =   4200
      Width           =   8775
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CL_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6000
      MaxLength       =   25
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CL_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4320
      MaxLength       =   25
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CL_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   7260
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
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
      Begin VB.CommandButton cmdViewReport 
         Appearance      =   0  'Flat
         Caption         =   "Employee Warning Report"
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   2205
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7560
         Top             =   105
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
         Height          =   330
         Left            =   8160
         Top             =   180
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
         Caption         =   ""
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   11055
      _Version        =   65536
      _ExtentX        =   19500
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
         Left            =   6600
         TabIndex        =   36
         Top             =   115
         Width           =   1305
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
         TabIndex        =   18
         Top             =   115
         Width           =   720
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
         TabIndex        =   17
         Top             =   110
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1005
      End
   End
   Begin INFOHR_Controls.DateLookup dlpEmpAgreedDate 
      DataField       =   "CL_EMP_AGREED"
      Height          =   285
      Left            =   7320
      TabIndex        =   6
      Tag             =   "40-Agreed Date"
      Top             =   2940
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpEmpDeclinedDate 
      DataField       =   "CL_EMP_DECLINED"
      Height          =   285
      Left            =   7320
      TabIndex        =   7
      Tag             =   "40-Declined Date"
      Top             =   3270
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpExpirationDate 
      DataField       =   "CL_EXPDATE"
      Height          =   285
      Left            =   7320
      TabIndex        =   11
      Tag             =   "41-Expiration Date"
      Top             =   6090
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSMask.MaskEdBox medLevel 
      DataField       =   "CL_LEVEL"
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Tag             =   "20-Level #"
      Top             =   6090
      Width           =   975
      _ExtentX        =   1720
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
      Format          =   "0"
      PromptChar      =   " "
   End
   Begin INFOHR_Controls.DateLookup dlpIncDate 
      DataField       =   "CL_INCDATE"
      Height          =   285
      Left            =   7320
      TabIndex        =   5
      Tag             =   "41-Incident Date"
      Top             =   2610
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Stream"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   40
      Top             =   6540
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Level #"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   120
      TabIndex        =   38
      Top             =   6135
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   5520
      TabIndex        =   37
      Top             =   6105
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Response"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   60
      TabIndex        =   35
      Top             =   5115
      Width           =   2070
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Declined Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   5520
      TabIndex        =   33
      Top             =   3285
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Agreed Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   5520
      TabIndex        =   32
      Top             =   2955
      Width           =   1815
   End
   Begin VB.Image imgNoSec 
      Height          =   240
      Left            =   7560
      Picture         =   "fECounsel.frx":3F80
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSec 
      Height          =   240
      Left            =   7560
      Picture         =   "fECounsel.frx":40CA
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblImport 
      Alignment       =   1  'Right Justify
      Caption         =   "Counseling"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   31
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Counseled By"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   29
      Top             =   3615
      Width           =   1155
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Counselling Date"
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
      Index           =   5
      Left            =   60
      TabIndex        =   28
      Top             =   2955
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   27
      Top             =   2625
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Index           =   1
      Left            =   60
      TabIndex        =   26
      Top             =   2625
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   25
      Top             =   3285
      Width           =   735
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "CL_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   7140
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EmpNbr"
      DataField       =   "CL_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1680
      TabIndex        =   23
      Top             =   7140
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   22
      Top             =   3915
      Width           =   2190
   End
End
Attribute VB_Name = "frmECounsel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode, fglbNew
Dim AddChg
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim OType, oCounselDt, OReason, OCounselBy, OIncidentDt, OEmpAgreeDt, OEmpDeclineDt, OComments, OEmpResp
Dim oExpirationDate

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdModify_Click()
    On Error GoTo Mod_Err
    
'    Call ST_UPD_MODE(True)
    Call SET_UP_MODE
    clpCode(1).SetFocus
    AddChg = "C"
    
    OType = clpCode(1).Text
    oCounselDt = dlpCouDate.Text
    OReason = clpCode(2).Text
    OCounselBy = elpCouByShow
    OIncidentDt = dlpIncDate.Text
    OEmpAgreeDt = dlpEmpAgreedDate.Text
    OEmpDeclineDt = dlpEmpDeclinedDate.Text
    OComments = memComments
    OEmpResp = memEmpResponse
    
    'Release 8.1
    oExpirationDate = dlpExpirationDate.Text
    
    fglbNew = False
    
    Exit Sub
    
Mod_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HR_COUNSEL", "Modify")
End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
    RHeading = lblEEName & "'s Assets"
Else
    RHeading = lblEEName & "'s " & lStr("Counseling")
End If
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        xReport = glbIHRREPORTS & "rgridcou_FN.rpt"
    Else
        xReport = glbIHRREPORTS & "rgridcou.rpt"
    End If
    vbxCrystal.ReportFileName = xReport
    vbxCrystal.SelectionFormula = "{HR_COUNSEL.CL_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    If glbSQL Or glbOracle Then
        vbxCrystal.Connect = RptODBC_SQL
    Else
        vbxCrystal.Connect = "PWD=petman;"
        vbxCrystal.DataFiles(0) = glbIHRAUDIT
        vbxCrystal.DataFiles(1) = glbIHRAUDIT
        vbxCrystal.DataFiles(2) = glbIHRDB
        vbxCrystal.DataFiles(3) = glbIHRDB
        vbxCrystal.DataFiles(4) = glbIHRDB
        vbxCrystal.DataFiles(5) = glbIHRDB
    End If
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        xReport = glbIHRREPORTS & "rgridcou_FN.rpt"
    Else
        xReport = glbIHRREPORTS & "rgridcouT.rpt"
    End If
    vbxCrystal.ReportFileName = xReport
    vbxCrystal.SelectionFormula = "{Term_HR_COUNSEL.TERM_SEQ}=" & glbTERM_Seq & " "
End If

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True
End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport, x%

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False

If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
    RHeading = lblEEName & "'s Assets"
Else
    RHeading = lblEEName & "'s " & lStr("Counseling")
End If
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        xReport = glbIHRREPORTS & "rgridcou_FN.rpt"
    Else
        xReport = glbIHRREPORTS & "rgridcou.rpt"
    End If
    vbxCrystal.ReportFileName = xReport
    vbxCrystal.SelectionFormula = "{HR_COUNSEL.CL_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    If glbSQL Or glbOracle Then
        vbxCrystal.Connect = RptODBC_SQL
    Else
        vbxCrystal.Connect = "PWD=petman;"
        vbxCrystal.DataFiles(0) = glbIHRDB
        vbxCrystal.DataFiles(1) = glbIHRDB
        vbxCrystal.DataFiles(2) = glbIHRDB
        vbxCrystal.DataFiles(3) = glbIHRDB
        vbxCrystal.DataFiles(4) = glbIHRAUDIT
        vbxCrystal.DataFiles(5) = glbIHRAUDIT
    End If
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        xReport = glbIHRREPORTS & "rgridcou_FN.rpt"
    Else
        xReport = glbIHRREPORTS & "rgridcouT.rpt"
    End If
    vbxCrystal.ReportFileName = xReport
    vbxCrystal.SelectionFormula = "{Term_HR_COUNSEL.TERM_SEQ}=" & glbTERM_Seq & " "
End If

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub

Private Sub clpCode_Change(Index As Integer)
'Ticket #24663 - Showa only
If glbCompSerial = "S/N - 2454W" And clpCode(1).Text = "COC" Then
    lblTitle(12).Visible = True
    'cmbStream.Visible = True
    txtStream.Visible = True
    cmdStreamRules.Visible = True
    
    'Ticket #25826 - For COC show the label caption from "Level #" to "Step #"
    lblTitle(11).Caption = "Step #"
    medLevel.Tag = "20-Step #"
Else
    lblTitle(12).Visible = False
    'cmbStream.Visible = False
    txtStream.Visible = False
    cmdStreamRules.Visible = False

    'Ticket #25826 - For non COC show the label caption from "Step #" to "Level #"
    lblTitle(11).Caption = "Level #"
    medLevel.Tag = "20-Level #"
End If

'Ticket #24663 - Showa only
If glbCompSerial = "S/N - 2454W" And (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC") Then
    lblTitle(2).FontBold = True
Else
    lblTitle(2).FontBold = False
End If

'Ticket #24663 - Showa only
'Compute Level # for new records with Type and Reason matching existing records
If glbCompSerial = "S/N - 2454W" And fglbNew Then
    If (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC") And Len(clpCode(2).Text) > 0 And IsDate(dlpIncDate.Text) Then
        'Compute the Level #. For ATT: 1 to 4, and for COC: 1 to 6
        Dim xLvl As Integer
        
        xLvl = Get_Level_Number(glbLEE_ID, clpCode(1).Text, clpCode(2).Text, dlpIncDate.Text)
        medLevel.Text = xLvl
        
        If clpCode(1).Text = "ATT" Then
            If xLvl > 4 Then
                MsgBox "For the matching Type and Reason, the Level # is exceeding 4. This record cannot be saved.", vbExclamation, "Level # Exceeding for Type ATT"
            End If
        ElseIf clpCode(1).Text = "COC" Then
            If xLvl > 6 Then
                MsgBox "For the matching Type and Reason, the Step # is exceeding 6. This record cannot be saved.", vbExclamation, "Step # Exceeding for Type COC"
            End If
        End If
    End If
End If

End Sub

Private Sub cmdStreamRules_Click()
Dim MsgStr As String
    
    Load frmMsgBoxList
    frmMsgBoxList.Caption = "Rules Implemented by Showa"
    frmMsgBoxList.lblQuestion.Caption = "The following rules have been implemented by Showa to promote and maintain a safe and efficient operation and to establish in advance what is expected of all Associates. Violation of any rule of conduct may result in progressive discipline up to and including immediate dismissal. These rules are meant only to give you an idea of the kinds of behaviour Showa considers inappropriate. The following list is not intended to be all-inclusive:"
    MsgStr = "1 - Fighting with or attempting to injure another Associate."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "2 - Possessing any weapon or related paraphernalia on SCI premises."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "3 - Creating an intimidating, hostile, or offensive working environment by threatening, coercing, retaliating  against or using abusive/threatening language towards Associates or visitors on SCI property."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "4 - Failing to immediately report a workplace illness, injury, incident or hazard to the Department Supervisor,  Manager or Health & Safety Specialist."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "5 - Failing to observe established safety rules including but not limited to, not wearing appropriate PPE as per operation standards."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "6 - Smoking or creating a flame outside the designated smoking area without authorization."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "7 - Engaging in illegal gambling, horseplay or practical jokes, including but not limited to any prank, contest, feat of strength, unnecessary running or rough and boisterous conduct or games of chance."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "8 - Possessing, using/misusing, distributing, selling, purchasing, reporting to work under the influence of, any illegal or prescription drugs/substances or alcohol while on SCI property."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "9 - Stealing, willfully damaging, abusing or hiding any property belonging to another Associates or SCI."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "10 - Giving false information with respect to absence, sickness, injury claims or personnel files."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "11 - Misrepresenting facts or falsifying records or reports, of a business nature."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "12 - Inappropriately obtaining or sharing confidential information with anyone outside of the appropriate SCI department."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "13 - Interfering, failing to cooperate, or divulging confidential information related to an authorized SCI investigation."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "14 - Obtaining property, money, or other privileges from SCI through fraud or misrepresentation or engaging in this type of activity while conducting SCI business."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "15 - Willfully scanning or requesting the scanning of yours or another Associate’s identification card."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "16 - Interfering with the work or performance of another associate or causing a restriction or slow down of production or tampering with or deliberately misusing company equipment."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "17 - Using electronic equipment, cameras, video equipment, or recording devices, without proper authorization or inappropriately while on SCI premises."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "18 - Using or carrying personal cell phones, personal paging devices or other electronic equipment on the floor unless otherwise authorized."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "19 - Reporting to the work area without wearing the proper SCI uniform including weekends, holidays and shutdowns."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "20 - Being absent from work for an unauthorized reason (any incident in which the associate fails to declare the absence as Personal Emergency Leave or the absence does not meet the qualifications under Personal Emergency Leave provisions)."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "21 - Failing to contact SCI to report an absence within the scheduled shift (No call / No show)."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "22 - Providing late notification of an absence (notification of less than 1 hour prior to shift start)."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "23 - Failing to swipe in prior to the scheduled shift start time (late)."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "24 - Leaving work prior to the end of the scheduled shift (Early Departure) without declaring it an Emergency Day."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "25 - Failing to provide appropriate medical documentation following an absence of 3 or more consecutive days."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "26 - Failing to complete the No Swipe Form and receive the Supervisor’s authorization (prior to the shift start meeting or at the end of shift). (Associates will be paid for the earliest time confirmed present at work.)"
    MsgStr = MsgStr & vbCrLf & vbCrLf & "27 - Reporting to the work area late and or leaving early at shift start up or after breaks and lunch or prior to shift end."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "28 - Leaving the work area during assigned working hours without permission."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "29 - Stopping work prior to the shift end."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "30 - Sleeping or loafing around while on the job."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "31 - Working below the operation standard for quality or quantity, willfully neglecting job responsibilities or refusing to comply with instructions of management."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "32 - Removing, defacing, or changing notices or bulletins posted throughout the facility or placing signs, notes, papers, or any materials not required for production in or on products or property."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "33 - Bringing or removing tools, materials or equipment to or from SCI premises without proper authorization."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "34 - Duplicating any SCI keys without the proper authorization."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "35 - Storing or posting inappropriate materials other than personal items and uniforms in SCI issued lockers."
    MsgStr = MsgStr & vbCrLf & vbCrLf & "36 - Parking inappropriately or outside the designated Associate parking area during working hours."
    frmMsgBoxList.txtLongMsg = MsgStr
    
    frmMsgBoxList.Show 1
    
    
    
End Sub

'Private Sub cmbStream_Change()
'If cmbStream.ListIndex <> -1 Then
'    txtStream.Text = Left(cmbStream.Text, InStr(1, cmbStream.Text, " ") - 1)
'Else
'    txtStream.Text = ""
'End If
'End Sub

'Private Sub cmbStream_Click()
'If cmbStream.ListIndex <> -1 Then
'    txtStream.Text = Left(cmbStream.Text, InStr(1, cmbStream.Text, " ") - 1)
'Else
'    txtStream.Text = ""
'End If
'End Sub

Private Sub cmdViewReport_Click()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim RHeading As String, xReport, x%
Dim SQLQ As String

RHeading = lblEEName & "'s Employee Warning"

Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

If Not glbtermopen Then
    Me.vbxCrystal.Connect = RptODBC_SQL
    xReport = glbIHRREPORTS & "RZCouWarn.rpt"
    vbxCrystal.ReportFileName = xReport
    
    TempCri = "({HR_COUNSEL.CL_COUDATE} "
    TempCri = TempCri & " = "
    dtYYY% = Year(dlpCouDate.Text)
    dtMM% = month(dlpCouDate.Text)
    dtDD% = Day(dlpCouDate.Text)
    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        
    
    SQLQ = "{HR_COUNSEL.CL_EMPNBR}=" & glbLEE_ID & " and {HR_COUNSEL.CL_TYPE}= '" & clpCode(1).Text & "' "
    SQLQ = SQLQ & "and " & TempCri
    
    vbxCrystal.SelectionFormula = SQLQ
    
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
End If


End Sub

Private Sub dlpIncDate_LostFocus()
    'Ticket #24663 - Showa only
    If glbCompSerial = "S/N - 2454W" And (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC") Then
        If IsDate(dlpIncDate.Text) And Not IsDate(dlpExpirationDate.Text) Then
            dlpExpirationDate.Text = DateAdd("m", 6, dlpIncDate.Text)
        ElseIf Not IsDate(dlpExpirationDate.Text) Then
            dlpExpirationDate.Text = ""
        End If
    End If
End Sub

Private Sub elpCouByShow_Change()
    txtCouBy.Text = getEmpnbr(elpCouByShow.Text)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMECOUNSEL"
AddChg = " "

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = vbDefault

'Ticket #24663 - Showa only
If glbCompSerial = "S/N - 2454W" Then
    'lblTitle(10).Visible = True    'Visible for all
    'lblTitle(11).Visible = True    'Visible for all
    'lblTitle(12).Visible = True    'Only if Type = COC
    'dlpExpirationDate.Visible = True   'Visible for all
    'medLevel.Visible = True            'Visible for all
    'cmbStream.Visible = True       'Only if Type = COC
End If

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_Counsel_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Counseling") & " information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Sub
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then 'Exit Sub
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
End If

Screen.MousePointer = vbHourglass

Me.vbxTrueGrid.SetFocus
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        Me.Caption = "Assets - " & Left$(glbLEE_SName, 5)
    Else
        Me.Caption = lStr("Counseling") & " - " & Left$(glbLEE_SName, 5)
    End If
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

lblEENum.Caption = ShowEmpnbr(lblEEID)

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'   cmdModify.Enabled = False
Else
'   cmdModify.Enabled = True
   Data1.Recordset.MoveFirst
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Call Display_Value

' TODO - Replace True with security check for Inquire/Maintain
'If True Then
Call ST_UPD_MODE(True)
'Else
'    Call ST_UPD_MODE(False)
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If

If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
    lblTitle(1) = "Type"
    lblTitle(2) = "Item"
    lblTitle(4) = "Returned Date"
    lblTitle(5) = "Issuing Date"
    lblTitle(6) = "Issued By"
    vbxTrueGrid.Columns(0).Caption = "Issuing Date"
    vbxTrueGrid.Columns(1).Caption = "Returned Date"
    'vbxTrueGrid.Columns(2).Caption = "Type"
    vbxTrueGrid.Columns(3).Caption = "Item"
    vbxTrueGrid.Columns(4).Caption = "Issued By"
    clpCode(1).TABLTitle = "Type Codes"
    clpCode(2).TABLTitle = "Item Codes"
End If

'Counseling labels
lblTitle(5).Caption = lStr("Counseling") & " Date"    'lStr(lblTitle(5).Caption)
lblTitle(6).Caption = lStr("Counseling") & " By"
lblImport.Caption = lStr(lblImport.Caption)
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
vbxTrueGrid.Columns(4).Caption = lStr("Counseling") & " By"

'Ticket #13036
If glbCompSerial = "S/N - 2382W" Then 'Namasco
    cmdViewReport.Visible = True
    lblTitle(3).Caption = "Supervisor Comments"
End If

End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_Counsel_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Counseling") & " information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_HR_COUNSEL"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * from HR_COUNSEL"
    SQLQ = SQLQ & " where CL_EMPNBR = " & glbLEE_ID
End If
'Order should be by CL_COUDATE, ticket #5832, Frank
'SQLQ = SQLQ & " ORDER BY CL_INCDATE DESC,CL_TYPE"
SQLQ = SQLQ & " ORDER BY CL_COUDATE DESC,CL_TYPE"
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Counsel Retrieve", "HR_COUNSEL", "SELECT")

Exit Function

End Function


Private Sub ST_UPD_MODE(YN)
Dim TF As Boolean, FT As Boolean

If YN Then TF = True Else TF = False
FT = Not TF

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
'memComments.Enabled = TF
memComments.Locked = FT
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
elpCouByShow.Enabled = TF
dlpIncDate.Enabled = TF
dlpCouDate.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'vbxTrueGrid.Enabled = FT

'Ticket #24663 - Showa
dlpExpirationDate.Enabled = TF
medLevel.Enabled = TF
If glbCompSerial = "S/N - 2454W" Then
    'Showa only
    'cmbStream.Enabled = TF
    txtStream.Enabled = TF
    cmdStreamRules.Enabled = TF
End If

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
'cmdPrint.Visible = Not glbtermopen
fUPMode = TF    ' update mode

'George on Jan 26,2006 #10266
glbDocName = "Counsel"
If gsAttachment_DB Then
    'glbJob = "" 'George on Jan 24,2006 #10266
    'glbSDate = "01/01/1900" 'George on Jan 24,2006 #10266
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        'glbJob = rsDATA("CL_TYPE") 'George on Jan 24,2006 #10266
        'glbSDate = rsDATA("CL_COUDATE") 'George on Jan 24,2006 #10266
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("CL_DOCKEY")) Then
                glbDocKey = rsDATA("CL_DOCKEY")
            Else
                If Not fglbNew Then
                    glbDocKey = 0
                End If
            End If
        Else
            If Not IsNull(Data1.Recordset("CL_DOCKEY")) Then
                glbDocKey = Data1.Recordset("CL_DOCKEY")
            Else
                If Not fglbNew Then
                    glbDocKey = 0
                End If
            End If
        End If
    End If
    'rsDATA
    
    Call DispimgIcon(Me, "frmECounsel")
    If gSec_Upd_Counselling Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If
'George on Jan 26,2006 #10266

End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

'George on Jan 26,2006 #10266
If gsAttachment_DB Then
    glbJob = ""
    glbSDate = "01/01/1900"
    lblImport.Visible = True 'False
    imgSec.Visible = False
    imgNoSec.Visible = True 'False
    cmdImport.Visible = True 'False
End If
'George on Jan 26,2006 #10266

clpCode(1).SetFocus

On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

glbCounselType = ""
glbCounselDate = ""

AddChg = "A"

OType = ""
oCounselDt = ""
OReason = ""
OCounselBy = ""
OIncidentDt = ""
OEmpAgreeDt = ""
OEmpDeclineDt = ""
OComments = ""
OEmpResp = ""

'Release 8.1
oExpirationDate = ""

fglbNew = True

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_COUNSEL", "Add")
End Sub

Sub cmdOK_Click()
Dim x, xID
Dim rsCOM As New ADODB.Recordset
On Error GoTo Add_Err

If Not chkECounsel() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

If fglbNew Then
    If Not AUDITCOUNSEL("A") Then MsgBox "ERROR - AUDIT FILE"
Else
    If Not AUDITCOUNSEL("M") Then MsgBox "ERROR - AUDIT FILE"
End If

rsDATA!CL_TYPE = clpCode(1).Text
rsDATA!CL_TYPE_TABL = "CETY"
rsDATA!CL_REASON_TABL = "CERE"

Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    rsDATA.Resync
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "UPDATE Term_HRDOC_COUNSEL SET DC_CLTYPE='" & rsDATA("CL_TYPE") & "',DC_COUDATE=" & Date_SQL(rsDATA("CL_COUDATE")) & " WHERE DC_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DC_DOCKEY=" & glbDocKey ' & " AND DC_COUDATE=" & Date_SQL(glbSDate) '" and DC_CLTYPE='" & glbJob &
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
    xID = rsDATA("CL_ID")
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    rsDATA.Resync
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "UPDATE HRDOC_COUNSEL SET DC_CLTYPE='" & rsDATA("CL_TYPE") & "',DC_COUDATE=" & Date_SQL(rsDATA("CL_COUDATE")) & " WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR = " & glbLEE_ID & " and DC_DOCKEY=" & glbDocKey ' & " AND DC_COUDATE=" & Date_SQL(glbSDate)   '" and DC_CLTYPE='" & glbJob &
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
    xID = rsDATA("CL_ID")
End If
Data1.Refresh
Data1.Recordset.Find "CL_ID=" & xID

'Call ST_UPD_MODE(True)

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            If glbtermopen Then
                Call AttachmentAdd(glbTERM_ID, glbDocImpFile, glbDocType, glbDocDesc)
            Else
                Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
            End If
        End If
    End If
    glbDocImpFile = ""
End If

'Release 8.1
If Not updFollow("U") Then Exit Sub

Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

If Not glbtermopen Then
    If glbWFC And glbPlantCode = "WHBY" Then
        Call Whitby60daysRule(glbLEE_ID, "D")
    End If
End If
        
fglbNew = False

Exit Sub

Add_Err:
If Err = 3022 Then
     MsgBox "Duplicate record existed - not entered"
     Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_COUNSEL", "Update")
End Sub

Private Function chkECounsel() As Boolean
    Dim rsEmp As New ADODB.Recordset
    Dim strCounselling As String
    Dim strIncident As String

    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        strCounselling = "Issuing"
        strIncident = "Returned"
    Else
        strCounselling = lStr("Counseling")
        strIncident = "Incident"
    End If
    If clpCode(1).Text = "" Then
        MsgBox "Type is a required field.", vbInformation + vbOKOnly, "Missing Information"
        clpCode(1).SetFocus
        Exit Function
    End If
    If dlpCouDate.Text = "" Then
        MsgBox strCounselling & " Date is a required field.", vbInformation + vbOKOnly, "Missing Information"
        dlpCouDate.SetFocus
        Exit Function
    End If
        If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Type code must be valid.", vbInformation + vbOKOnly, "Missing Information"
        clpCode(1).SetFocus
        Exit Function
    End If
    If clpCode(2).Caption = "Unassigned" And clpCode(2).Text <> "" Then
        If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
            MsgBox "If entered, Item code must be valid.", vbInformation + vbOKOnly, "Missing Information"
        Else
            MsgBox "If entered, Reason code must be valid.", vbInformation + vbOKOnly, "Missing Information"
        End If
        clpCode(2).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpCouDate.Text) Then
        MsgBox strCounselling & " Date is not a valid date.", vbInformation + vbOKOnly, "Missing Information"
        dlpCouDate.SetFocus
        Exit Function
    End If
    
      'Hemu 05/09/2003 Begin - Counselling Date and Original Hire Date
    
    
    If glbtermopen Then
        rsEmp.Open "SELECT ED_DOH,ED_SENDTE FROM Term_HREMP WHERE ED_EMPNBR = " & lblEENum, gdbAdoIhr001X, adOpenStatic
    Else
        rsEmp.Open "SELECT ED_DOH,ED_SENDTE FROM HREMP WHERE ED_EMPNBR = " & lblEENum, gdbAdoIhr001, adOpenStatic
    End If
    
    If Not rsEmp.EOF Then
        If glbSamuel Then 'Ticket #23523 Franks 04/05/2013
            If Not IsNull(rsEmp("ED_SENDTE")) Then
                If DaysBetween(rsEmp("ED_SENDTE"), dlpCouDate.Text) < 0 Then
                    MsgBox strCounselling & " Date can not be prior to the Employee " & lStr("Seniority") & " date"
                    dlpCouDate.SetFocus
                    rsEmp.Close
                    Exit Function
                End If
            End If
        Else
            If Not IsNull(rsEmp("ED_DOH")) Then
                If DaysBetween(rsEmp("ED_DOH"), dlpCouDate.Text) < 0 Then
                    MsgBox strCounselling & " Date can not be prior to Original Hire date"
                    dlpCouDate.SetFocus
                    rsEmp.Close
                    Exit Function
                End If
            End If
        End If
    End If
    rsEmp.Close
    'Hemu 05/09/2003 End
    
    If (glbWFC And glbPlantCode = "WHBY") Or (glbCompSerial = "S/N - 2454W" And (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC")) Then
        If Len(dlpIncDate.Text) = 0 Then
            MsgBox strIncident & " Date is required."
            dlpIncDate.SetFocus
            Exit Function
        End If
    End If
    
    If Not IsDate(dlpIncDate.Text) And dlpIncDate.Text <> "" Then
        MsgBox strIncident & " Date is not a valid date.", vbInformation + vbOKOnly, "Missing Information"
        dlpIncDate.SetFocus
        Exit Function
    End If
    
     'Hemu 05/09/2003 Begin - Incident Date and Counselling Date
    If Len(dlpIncDate.Text) > 0 Then
        If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
            If DaysBetween(dlpCouDate.Text, dlpIncDate.Text) < 0 Then
                MsgBox strIncident & " Date must be greater than " & strCounselling & " Date"
                dlpIncDate.SetFocus
                Exit Function
            End If
        Else
            If DaysBetween(dlpIncDate.Text, dlpCouDate.Text) < 0 Then
                MsgBox strIncident & " Date can not be greater than " & strCounselling & " Date"
                dlpIncDate.SetFocus
                Exit Function
            End If
        End If
    End If
    'Hemu 05/09/2003
    If elpCouByShow.Text <> "" And (elpCouByShow.Caption = "Unassigned" Or elpCouByShow.Caption = "Enter Valid Employee #") Then
        If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
            MsgBox "If entered, Issued By must be a valid employee number.", vbInformation + vbOKOnly, "Missing Information"
        Else
            MsgBox "If entered, Counselled By must be a valid employee number.", vbInformation + vbOKOnly, "Missing Information"
        End If
        elpCouByShow.SetFocus
        Exit Function
    End If
    
    
    If Len(dlpEmpAgreedDate.Text) > 0 And Len(dlpEmpDeclinedDate.Text) > 0 Then
            MsgBox "Can not enter both Employee Agreed Date and Employee Declined Date"
            dlpEmpDeclinedDate.SetFocus
            Exit Function
    End If
    If Not IsDate(dlpEmpAgreedDate.Text) And dlpEmpAgreedDate.Text <> "" Then
        MsgBox "Employee Agreed Date is not a valid date.", vbInformation + vbOKOnly, "Missing Information"
        dlpEmpAgreedDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpEmpDeclinedDate.Text) And dlpEmpDeclinedDate.Text <> "" Then
        MsgBox "Employee Declined Date is not a valid date.", vbInformation + vbOKOnly, "Missing Information"
        dlpEmpDeclinedDate.SetFocus
        Exit Function
    End If
    
    'Ticket #24663 - Showa only
    'Reason is required for the COC and ATT Type because it is being used to compute the Level #
    If glbCompSerial = "S/N - 2454W" Then
        If (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC") And Len(Trim(clpCode(2).Text)) = 0 Then
            MsgBox "Reason is a required when Type is 'ATT' or 'COC'."
            clpCode(2).SetFocus
            Exit Function
        End If
    End If
    
    'Ticket #24663 - Showa only
    'Validate or Compute the Level # for new records with Type and Reason matching existing records
    If glbCompSerial = "S/N - 2454W" And fglbNew Then
        If (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC") And Len(clpCode(2).Text) > 0 And IsDate(dlpIncDate.Text) Then
            If IsNumeric(medLevel) Then
                'Validate the Level #
                If clpCode(1).Text = "ATT" Then
                    If medLevel.Text > 4 Then
                        MsgBox "Level # is exceeding 4. This record cannot be saved.", vbExclamation, "Level # Exceeding for Type ATT"
                        Exit Function
                    End If
                ElseIf clpCode(1).Text = "COC" Then
                    If medLevel.Text > 6 Then
                        MsgBox "Step # is exceeding 6. This record cannot be saved.", vbExclamation, "Step # Exceeding for Type COC"
                        Exit Function
                    End If
                End If
            Else
                'Compute the Level #. For ATT: 1 to 4, and for COC: 1 to 6
                Dim xLvl As Integer
                
                xLvl = Get_Level_Number(glbLEE_ID, clpCode(1).Text, clpCode(2).Text, dlpIncDate.Text)
                medLevel.Text = xLvl
                
                If clpCode(1).Text = "ATT" Then
                    If xLvl > 4 Then
                        MsgBox "For the matching Type and Reason, the Level # is exceeding 4. This record cannot be saved.", vbExclamation, "Level # Exceeding for Type ATT"
                        Exit Function
                    End If
                ElseIf clpCode(1).Text = "COC" Then
                    If xLvl > 6 Then
                        MsgBox "For the matching Type and Reason, the Step # is exceeding 6. This record cannot be saved.", vbExclamation, "Step # Exceeding for Type COC"
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    'Ticket #24663 - Showa only
    'Check for duplicates
    If glbCompSerial = "S/N - 2454W" Then
        If (clpCode(1).Text = "ATT" Or clpCode(1).Text = "COC") And Len(clpCode(2).Text) > 0 Then
            Dim flgDuplicate As Boolean
            flgDuplicate = False
            
            If fglbNew Then
                flgDuplicate = Duplicate_Counselling_Level(glbLEE_ID, clpCode(1).Text, clpCode(2).Text, medLevel.Text, "New", , dlpIncDate.Text)
            Else
                flgDuplicate = Duplicate_Counselling_Level(glbLEE_ID, clpCode(1).Text, clpCode(2).Text, medLevel.Text, "Modify", Data1.Recordset!CL_ID, dlpIncDate.Text)
            End If
            
            If flgDuplicate Then
                If clpCode(1).Text = "ATT" Then
                    MsgBox "Duplicate record matching the Type, Reason and the Level #. This record cannot be saved.", vbExclamation, "Duplicate record"
                ElseIf clpCode(1).Text = "COC" Then
                    MsgBox "Duplicate record matching the Type, Reason and the Step #. This record cannot be saved.", vbExclamation, "Duplicate record"
                End If
                Exit Function
            End If
            
            If clpCode(1).Text = "COC" And (IsNumeric(txtStream.Text) Or Len(txtStream.Text) > 0) Then
                If Val(txtStream.Text) < 1 Or Val(txtStream.Text) > 36 Then
                    MsgBox "Invalid Stream"
                    txtStream.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
    
    chkECounsel = True
End Function

Sub cmdCancel_Click()
    Dim x
    On Error GoTo Can_Err
    
    rsDATA.CancelUpdate
    
    fglbNew = False
    
    Call Display_Value
    
    fglbNew = False
    
    Call SET_UP_MODE
    'Call ST_UPD_MODE(True)  ' reset screen's attributes
    
    Me.vbxTrueGrid.SetFocus
    
    Exit Sub
    
Can_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cancel", "HR_COUNSEL", "Cancel")
End Sub

Sub cmdDelete_Click()
    Dim a As Integer, Msg As String, x
    
    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    
    On Error GoTo Del_Err
    
    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & "This Record?"
    a% = MsgBox(Msg, 36, "Confirm Delete")
    
    If a% <> 6 Then Exit Sub
    
    oExpirationDate = dlpExpirationDate.Text
    If oExpirationDate <> "" Then
        If Not updFollow("D") Then
            Exit Sub
        End If
    End If
    
    If Not AUDITCOUNSEL("D") Then MsgBox "ERROR - AUDIT FILE"
    
    If glbtermopen Then
        'gdbAdoIhr001X.BeginTrans
        rsDATA.Delete
        'gdbAdoIhr001X.CommitTrans
        'George Jan 26,2006
        If gsAttachment_DB Then
            gdbAdoIhr001_DOC.BeginTrans
            'gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_COUNSEL where DC_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DC_CLTYPE='" & glbJob & "' and DC_COUDATE=" & Date_SQL(glbSDate)
            gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_COUNSEL where DC_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DC_DOCKEY=" & glbDocKey & " " '
            gdbAdoIhr001_DOC.CommitTrans
        End If
        'George Jan 26,2006
        Data1.Refresh
    Else
        If glbWFC And glbPlantCode = "WHBY" Then
            Call WhitbyDisciplineDelete(glbLEE_ID, clpCode(1), dlpIncDate)
        End If
        'Don't delete the Attendance record
        'If glbBurlTech Then
        '    Call BurlTechDisciplineDelete(glbLEE_ID, clpCode(1), dlpIncDate)
        'End If
        'gdbAdoIhr001.BeginTrans
        rsDATA.Delete
        'gdbAdoIhr001.CommitTrans
        'George Jan 26,2006
        If gsAttachment_DB Then
            'gdbAdoIhr001_DOC.BeginTrans 'glbDocKey = rsDATA("CL_DOCKEY")
            'gdbAdoIhr001_DOC.Execute "delete from HRDOC_COUNSEL where DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR = " & glbLEE_ID & " and DC_CLTYPE='" & glbJob & "' and DC_COUDATE=" & Date_SQL(glbSDate)
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_COUNSEL where DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR = " & glbLEE_ID & " and DC_DOCKEY=" & glbDocKey & " " '' and DC_COUDATE=" & Date_SQL(glbSDate)
            'gdbAdoIhr001_DOC.CommitTrans
        End If
        'George Jan 26,2006
        If glbWFC And glbPlantCode = "WHBY" Then
            Call Whitby60daysRule(glbLEE_ID, "D")
        End If
        If Not glbSQL And Not glbOracle Then
            Call Pause(0.5)
        End If
        Data1.Refresh
    End If
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Call Display_Value
    End If
    
    fglbNew = False
    
    'Call ST_UPD_MODE(True)
    Call SET_UP_MODE
    Exit Sub
    
Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HR_COUNSEL", "Delete")
End Sub

Private Sub BurlTechDisciplineDelete(xEmpNo, xType, xIncDate)
    'If Disciplinary action was deleted, the matched Attendance record should be delete
    Dim rsTemp As New ADODB.Recordset
    Dim rsTem2 As New ADODB.Recordset
    Dim SQLQ, xDiscipStep, xNextStepPlus, xNextStepVal

    ''Disable it until they are ready
    'Exit Sub
    
    'Delete the matched Attendance record
    SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & " AND AD_DISCIPLINE = '" & xType & "'"
    SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(xIncDate) & " "
    gdbAdoIhr001.Execute SQLQ
    Unload frmVATTEND
End Sub

Private Sub WhitbyDisciplineDelete(xEmpNo, xType, xIncDate)
    'If Disciplinary action was deleted, the matched Attendance record should be delete
    'Also check ED_DISCIPLINENEXT in HREMP, if this Disciplinary action is the current action
    'then ED_DISCIPLINENEXT should be ED_DISCIPLINENEXT -1
    Dim rsTemp As New ADODB.Recordset
    Dim rsTem2 As New ADODB.Recordset
    Dim SQLQ, xDiscipStep, xNextStepPlus, xNextStepVal

    ''Disable it until Whitby is ready
    'Exit Sub
    
    If Not IsDate(xIncDate) Then Exit Sub
    SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
    SQLQ = SQLQ & "AND CL_INCDATE = " & Date_SQL(xIncDate) & " "
    SQLQ = SQLQ & "AND CL_LDATE >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
    
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF Then Exit Sub
    'If IsNull(rsTemp("CL_ATTDATE")) Then Exit Sub
    'If Not IsDate(rsTemp("CL_ATTDATE")) Then Exit Sub
    'If IsNull(rsTemp("CL_ATTREASON")) Then Exit Sub
    'If Len(Trim(rsTemp("CL_ATTREASON"))) = 0 Then Exit Sub
    
    'Check if xType in HR_DISCIPLINE_STEPS table
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_DISCIPLINE = '" & xType & "' "
    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTem2.EOF Then
        rsTem2.Close
        Exit Sub
    Else
        xDiscipStep = rsTem2("DS_STEPNO")
    End If
    rsTem2.Close
    
    'Delete the matched Attendance record
    If Not IsNull(rsTemp("CL_ATTREASON")) Then
        SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AD_REASON = '" & rsTemp("CL_ATTREASON") & "'"
        SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(rsTemp("CL_ATTDATE")) & " "
        gdbAdoIhr001.Execute SQLQ
        Unload frmVATTEND
    End If
    rsTemp.Close
    

''    'Check if this Disciplinary action is the current action
''    SQLQ = "SELECT ED_EMPNBR, ED_DISCIPLINENEXT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
''    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''    xNextStepVal = 1
''    If Not rsTemp.EOF Then
''        If Not IsNull(rsTemp("ED_DISCIPLINENEXT")) Then
''            xNextStepVal = rsTemp("ED_DISCIPLINENEXT")
''            'If (rsTemp("ED_DISCIPLINENEXT") = xDiscipStep + 1) And xDiscipStep >= 1 Then
''            '    rsTemp("ED_DISCIPLINENEXT") = xDiscipStep
''            '    rsTemp.Update
''            'End If
''        End If
''    End If
''    rsTemp.Close
''
''    If Not (xDiscipStep + 1 = xNextStepVal) Then
''        Exit Sub
''    Else
''        'Check if this Counsel record is current Disciplinary step, but there are multi current steps
''        'if this Counsel record is not the latest record, don't change ED_DISCIPLINENEXT
''        SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
''        SQLQ = SQLQ & "AND CL_TYPE = '" & xType & "' "
''        SQLQ = SQLQ & "AND CL_LDATE >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
''        SQLQ = SQLQ & "ORDER BY CL_INCDATE DESC"
''        If rsTem2.State <> 0 Then rsTem2.Close
''        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
''        If Not rsTem2.EOF Then
''            If Not (CVDate(rsTem2("CL_INCDATE")) = CVDate(xIncDate)) Then
''                rsTem2.Close
''                Exit Sub
''            End If
''        End If
''        rsTem2.Close
''    End If
''
''    SQLQ = "UPDATE HREMP SET ED_DISCIPLINENEXT = " & WhitbyPreStep(xEmpNo, xDiscipStep) & " WHERE ED_EMPNBR = " & xEmpNo
''    gdbAdoIhr001.Execute SQLQ
''
''    If xDiscipStep > 0 Then
''        'Reset the current Disciplinary
''        'To false
''        SQLQ = "UPDATE HR_COUNSEL SET CL_COMPLETED = 0 WHERE CL_EMPNBR = " & xEmpNo & " "
''        SQLQ = SQLQ & "AND CL_COMPLETED <> 0 "
''        gdbAdoIhr001.Execute SQLQ
''        'Get current Disciplinary Code
''        SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_STEPNO = " & xDiscipStep - 1 & " "
''        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
''        If Not rsTem2.EOF Then
''            SQLQ = "UPDATE HR_COUNSEL SET CL_COMPLETED = -1 WHERE CL_EMPNBR = " & xEmpNo & " "
''            SQLQ = SQLQ & "AND CL_TYPE = '" & rsTem2("DS_DISCIPLINE") & "' "
''            gdbAdoIhr001.Execute SQLQ
''            'SQLQ = "SELECT * FROM HR_COUNSEL WHERE CL_EMPNBR = " & xEmpNo & " "
''            'SQLQ = SQLQ & "AND CL_TYPE = '" & rsTem2("DS_DISCIPLINE") & "' "
''            'SQLQ = SQLQ & "ORDER BY CL_INCDATE DESC "
''            'rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''            'If Not rsTemp.EOF Then
''            '    rsTemp("CL_COMPLETED") = -1
''            '    rsTemp.Update
''            'End If
''            'rsTemp.Close
''        End If
''        rsTem2.Close
''    End If
End Sub

Function WhitbyPreStep(xEmpNo, xStep)
Dim rsDisci As New ADODB.Recordset
Dim rsCounsel As New ADODB.Recordset
Dim SQLQ, xPreStep, I, xMum
    xPreStep = xStep
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_STEPNO <= " & xStep & " ORDER BY DS_STEPNO DESC"
    rsDisci.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsDisci.EOF Then
        rsDisci.Close
        GoTo End_line
    End If
    xPreStep = 1
    Do While Not rsDisci.EOF
        SQLQ = "SELECT CL_EMPNBR FROM HR_COUNSEL WHERE CL_TYPE = '" & rsDisci("DS_DISCIPLINE") & "' "
        SQLQ = SQLQ & "AND  CL_EMPNBR = " & xEmpNo
        If rsCounsel.State <> 0 Then rsCounsel.Close
        rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsCounsel.EOF Then
            xPreStep = rsDisci("DS_STEPNO")
            rsCounsel.Close
            rsDisci.Close
            GoTo End_line
        End If
        rsCounsel.Close
        rsDisci.MoveNext
    Loop
    
End_line:
    WhitbyPreStep = xPreStep
End Function

''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Exit Sub
    End If
        
    If glbtermopen Then
        SQLQ = "Select * from Term_HR_COUNSEL"
        SQLQ = SQLQ & " WHERE CL_ID = " & Data1.Recordset!CL_ID
        SQLQ = SQLQ & " ORDER BY CL_COUDATE DESC,CL_TYPE"
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        
    Else
        SQLQ = "Select * from HR_COUNSEL"
        SQLQ = SQLQ & " where CL_ID = " & Data1.Recordset!CL_ID
        SQLQ = SQLQ & " ORDER BY CL_COUDATE DESC,CL_TYPE"
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
    Me.cmdModify_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Keepfocus As Boolean
    
    If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    Keepfocus = Not isUpdated(Me)
    Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCouBy_Change()
    elpCouByShow.Text = ShowEmpnbr(txtCouBy.Text)
End Sub

Private Sub txtStream_Change()
'If Len(txtStream.Text) > 0 Then
'    If IsNumeric(txtStream.Text) Then
'        cmbStream.ListIndex = CInt(txtStream.Text) - 1
'    Else
'        cmbStream.ListIndex = -1
'    End If
'Else
'    cmbStream.ListIndex = -1
'End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    If glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "Select * from Term_HR_COUNSEL"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "Select * from HR_COUNSEL"
        SQLQ = SQLQ & " where CL_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
'   Set FRS = Data1.Recordset.Clone
'   vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

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
UpdateRight = gSec_Upd_Counselling 'gSec_Upd_Basic
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
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
        frmECounsel.Caption = "Assets - " & Left$(glbLEE_SName, 5)
    Else
        frmECounsel.Caption = lStr("Counseling") & " - " & Left$(glbLEE_SName, 5)
    End If
    frmECounsel.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmECounsel")
    Call FillMemoFile(SQLQ, "Counsel")
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "Counsel"
    
    'Ticket #28839
    If fglbNew Then
        glbDocKey = 0
    Else
        'glbDocKey = rsDATA("CL_ID")
        If Not IsNull(rsDATA("CL_DOCKEY")) Then
            glbDocKey = rsDATA("CL_DOCKEY")
        Else
            glbDocKey = rsDATA("CL_ID") 'Ticket #16018
        End If
        glbCounselType = rsDATA("CL_TYPE")
        glbCounselDate = rsDATA("CL_COUDATE")
    End If
    
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmECounsel")
End Sub

Private Function AUDITCOUNSEL(ACTX)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITCOUNSEL = False

'rsTB.Open "HREMP", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
'rsTB.Find "ED_EMPNBR = " & glbLEE_ID
rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If

strFields = "AU_TYPE_TABL, AU_REASON_TABL, AU_ATTREASON_TABL, "
strFields = strFields & "AU_PTUPL, AU_DIVUPL, AU_COMMENTS, AU_COUBY, AU_COMPLETED, AU_EMP_RESPONSE, "
strFields = strFields & "AU_TYPE, AU_REASON, AU_ATTREASON, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TRANS_TYPE, "
strFields = strFields & "AU_COUDATE, AU_INCDATE, AU_FOLLOWUPD1, AU_FOLLOWUPD2, AU_FOLLOWUPD3, AU_ATTDATE, AU_DATE1, AU_EMP_AGREED, AU_EMP_DECLINED "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT_COUNSEL WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

If ACTX = "D" Or ACTX = "A" Then GoTo MODUPD
If OType <> clpCode(1).Text Then GoTo MODUPD
If oCounselDt <> dlpCouDate.Text Then GoTo MODUPD
If OReason <> clpCode(2).Text Then GoTo MODUPD
If OCounselBy <> elpCouByShow Then GoTo MODUPD
If OIncidentDt <> dlpIncDate.Text Then GoTo MODUPD
If OEmpAgreeDt <> dlpEmpAgreedDate.Text Then GoTo MODUPD
If OEmpDeclineDt <> dlpEmpDeclinedDate.Text Then GoTo MODUPD
If OComments <> memComments.Text Then GoTo MODUPD
If OEmpResp <> memEmpResponse.Text Then GoTo MODUPD

GoTo MODNOUPD

MODUPD:
rsTA.AddNew
rsTA("AU_TYPE_TABL") = "CETY": rsTA("AU_REASON_TABL") = "CERE": rsTA("AU_ATTREASON_TABL") = "ADRE"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If ACTX = "D" Then
    rsTA("AU_TYPE") = clpCode(1).Text
    rsTA("AU_COUDATE") = dlpCouDate.Text
Else
    rsTA("AU_TYPE") = clpCode(1).Text
    rsTA("AU_COUDATE") = dlpCouDate.Text
    
    If OReason <> clpCode(2).Text Then
        If clpCode(2).Text <> "" Then rsTA("AU_REASON") = clpCode(2).Text
    End If
    If OCounselBy <> elpCouByShow Then
        'If elpCouByShow.Text <> "" Then rsTA("AU_COUBY") = elpCouByShow.Text
        If elpCouByShow.Text <> "" Then rsTA("AU_COUBY") = getEmpnbr(elpCouByShow.Text)
    End If
    If OIncidentDt <> dlpIncDate Then
        If IsDate(dlpIncDate.Text) Then rsTA("AU_INCDATE") = dlpIncDate.Text
    End If
    If OEmpAgreeDt <> dlpEmpAgreedDate Then
        If IsDate(dlpEmpAgreedDate.Text) Then rsTA("AU_EMP_AGREED") = dlpEmpAgreedDate.Text
    End If
    If OEmpDeclineDt <> dlpEmpDeclinedDate Then
        If IsDate(dlpEmpDeclinedDate.Text) Then rsTA("AU_EMP_DECLINED") = dlpEmpDeclinedDate.Text
    End If
    If OComments <> memComments Then
        If memComments.Text <> "" Then rsTA("AU_COMMENTS") = memComments.Text
    End If
    If OEmpResp <> memEmpResponse Then
        If memEmpResponse.Text <> "" Then rsTA("AU_EMP_RESPONSE") = memEmpResponse.Text
    End If
End If

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TRANS_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITCOUNSEL = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '26July99 js

End Function

Private Function Get_Level_Number(xEmpNo, xType, xReason, xIncDate)
    Dim rsCounsel As New ADODB.Recordset
    Dim SQLQ As String

    'Level # for ATT: 1 to 4; and for COC: 1 to 6
    
    SQLQ = "SELECT CL_EMPNBR, CL_TYPE, CL_REASON, CL_LEVEL, CL_INCDATE, CL_EXPDATE FROM HR_COUNSEL WHERE CL_EMPNBR= " & xEmpNo
    SQLQ = SQLQ & " AND CL_TYPE = '" & xType & "'"
    SQLQ = SQLQ & " AND CL_REASON = '" & xReason & "'"
    SQLQ = SQLQ & " ORDER BY CL_LEVEL DESC, CL_EXPDATE DESC"
    rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsCounsel.EOF Then
        rsCounsel.MoveFirst
        
        Do While Not rsCounsel.EOF
            If IsNumeric(rsCounsel("CL_LEVEL")) Then
                If IsDate(rsCounsel("CL_EXPDATE")) Then
                    If CVDate(xIncDate) < CVDate(rsCounsel("CL_EXPDATE")) Then
                        'Within 6 months - move to next level
                        Get_Level_Number = rsCounsel("CL_LEVEL") + 1
                        Exit Do
                    'ElseIf CVDate(xIncDate) >= CVDate(DateAdd("m", "6", rsCounsel("CL_EXPDATE"))) And CVDate(xIncDate) < CVDate(DateAdd("m", "12", rsCounsel("CL_EXPDATE"))) Then
                    ElseIf CVDate(xIncDate) >= CVDate(rsCounsel("CL_EXPDATE")) And CVDate(xIncDate) < CVDate(DateAdd("m", "12", rsCounsel("CL_EXPDATE"))) Then
                        '6 to 12 months - re-issue same Level
                        Get_Level_Number = rsCounsel("CL_LEVEL")
                        Exit Do
                    ElseIf CVDate(xIncDate) >= CVDate(DateAdd("m", "12", rsCounsel("CL_EXPDATE"))) And CVDate(xIncDate) < CVDate(DateAdd("m", "18", rsCounsel("CL_EXPDATE"))) Then
                        '12 to 18 months - re-issue same Level minus 1
                        Get_Level_Number = rsCounsel("CL_LEVEL") - 1
                        
                        If Get_Level_Number <= 0 Then Get_Level_Number = 1
                        
                        Exit Do
                    ElseIf CVDate(xIncDate) >= CVDate(DateAdd("m", "18", rsCounsel("CL_EXPDATE"))) And CVDate(xIncDate) < CVDate(DateAdd("m", "24", rsCounsel("CL_EXPDATE"))) Then
                        '18 to 24 months - re-issue same Level minus 2
                        Get_Level_Number = rsCounsel("CL_LEVEL") - 2
                        
                        If Get_Level_Number <= 0 Then Get_Level_Number = 1
                        
                        Exit Do
                    ElseIf CVDate(xIncDate) >= CVDate(DateAdd("m", "24", rsCounsel("CL_EXPDATE"))) And CVDate(xIncDate) < CVDate(DateAdd("m", "30", rsCounsel("CL_EXPDATE"))) Then
                        '24 to 30 months - re-issue same Level minus 3
                        Get_Level_Number = rsCounsel("CL_LEVEL") - 3
                        
                        If Get_Level_Number <= 0 Then Get_Level_Number = 1
                        
                        Exit Do
                    ElseIf CVDate(xIncDate) >= CVDate(DateAdd("m", "30", rsCounsel("CL_EXPDATE"))) And CVDate(xIncDate) < CVDate(DateAdd("m", "36", rsCounsel("CL_EXPDATE"))) Then
                        'ATT
                        '30 to 36 months - start from Level 1
                        If xType = "ATT" Then
                            Get_Level_Number = 1
                        ElseIf xType = "COC" Then
                            'COC
                            '30 to 36 months - re-issue same Level minus 4
                            Get_Level_Number = rsCounsel("CL_LEVEL") - 4
                            
                            If Get_Level_Number <= 0 Then Get_Level_Number = 1
                        End If
                        Exit Do
                    ElseIf CVDate(xIncDate) >= CVDate(DateAdd("m", "36", rsCounsel("CL_EXPDATE"))) Then ' And CVDate(xIncDate) < CVDate(DateAdd("m", "36", rsCounsel("CL_EXPDATE"))) Then
                        '>= 36 months - start from Level 1
                        Get_Level_Number = 1
                        Exit Do
                    End If
                End If
            Else
                Get_Level_Number = 1
            End If
            
            rsCounsel.MoveNext
        Loop
    Else
        Get_Level_Number = 1
    End If
    rsCounsel.Close
    Set rsCounsel = Nothing

End Function

Private Function Duplicate_Counselling_Level(xEmpNo, xType, xReason, xLevel, xMode, Optional xID, Optional xIncDate) As Boolean
    Dim rsCounsel As New ADODB.Recordset
    Dim SQLQ As String

    Duplicate_Counselling_Level = False
    
    SQLQ = "SELECT CL_EMPNBR, CL_TYPE, CL_REASON, CL_LEVEL FROM HR_COUNSEL WHERE CL_EMPNBR= " & xEmpNo
    SQLQ = SQLQ & " AND CL_TYPE = '" & xType & "'"
    SQLQ = SQLQ & " AND CL_REASON = '" & xReason & "'"
    SQLQ = SQLQ & " AND CL_LEVEL = '" & xLevel & "'"
    SQLQ = SQLQ & " AND CL_INCDATE = " & Date_SQL(xIncDate)
    If xMode = "Modify" And Not IsMissing(xID) Then
        SQLQ = SQLQ & " AND CL_ID <> " & xID
    End If
    rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not rsCounsel.EOF Then
        Duplicate_Counselling_Level = True
    Else
        Duplicate_Counselling_Level = False
    End If
    rsCounsel.Close
    Set rsCounsel = Nothing

End Function

Private Function updFollow(xType)
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim rsTT As New ADODB.Recordset
Dim Edit1 As Integer
Dim ODate, xDATE

'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

ODate = oExpirationDate
xDATE = dlpExpirationDate.Text


On Error GoTo CrFollow_Err

If IsDate(ODate) Then     'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'COUN'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(ODate)
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    'New record and date is entered -> create follow up
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If fglbNew And IsDate(xDATE) Then
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'COUN'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xDATE)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
            'Create the Code if not already existing
            rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='COUN'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                rsTT("TB_KEY") = "COUN"
                rsTT("TB_DESC") = "Counselling Expiration"
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            Set rsTT = Nothing
            
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            Call Grant_FollowUpCode_Security(glbUserID, "COUN", "Counselling Expiration")
            
            'Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(xDATE)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = "COUN"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
    
    'Updating existing record but Follow Up record do not exists and the Date is valid -> create Follow Up
    If fglbNew = False And Edit1 = False And IsDate(xDATE) Then
        ' 5/2/2001 Add by Frank for no duplicated record of HR_FOLLOW_UP Begin
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'COUN' "
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xDATE)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
            'Create the Code if not already existing
            rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='COUN'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                rsTT("TB_KEY") = "COUN"
                rsTT("TB_DESC") = "Counselling Expiration"
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            Set rsTT = Nothing
        
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            Call Grant_FollowUpCode_Security(glbUserID, "COUN", "Counselling Expiration")
        
            'Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(xDATE)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = "COUN"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
  
    'Updating existing record and Follow Up record found, the Date is valid -> update existing Follow Up record
    If fglbNew = False And Edit1 = True And IsDate(xDATE) Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(xDATE)
            dynHRAT("EF_FREAS") = "COUN"
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If ODate <> xDATE Then
            'Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    
    'Updating existing record and the Follow Up exist, the Date is not valid -> delete the Follow Up record
    If fglbNew = False And Edit1 = True And Not IsDate(xDATE) Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
       ' Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If xDATE = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Counselling Expiration", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function

