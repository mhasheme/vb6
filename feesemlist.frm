VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmESEMList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Completed Course List"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7260
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCourseName 
      Appearance      =   0  'Flat
      DataField       =   "ES_COURSE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2300
      MaxLength       =   125
      TabIndex        =   5
      Tag             =   "01-Course Name"
      Top             =   3680
      Width           =   3855
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feesemlist.frx":0000
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "feesemlist.frx":0014
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   4335
      Width           =   7260
      _Version        =   65536
      _ExtentX        =   12806
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   900
         TabIndex        =   7
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   75
         TabIndex        =   2
         Tag             =   "Select Ledger Description"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6225
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "General Leager Codes and Descriptions"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   3720
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CRSCODE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   1980
      TabIndex        =   4
      Tag             =   "00-Course Code"
      Top             =   3330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
      MaxLength       =   8
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ES_CTYPE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   3
      Tag             =   "01-Course Type - Code"
      Top             =   3000
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCT"
      MaxLength       =   8
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpDatComp 
      DataField       =   "ES_DATCOMP"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Tag             =   "41-Date when course was completed"
      Top             =   4020
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   3680
      Width           =   1380
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Completed"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   4020
      Width           =   1725
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3030
      Width           =   1680
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3330
      Width           =   1680
   End
End
Attribute VB_Name = "frmESEMList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLQ As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSelect_Click()
Call SaveDataInArray
Unload Me
End Sub

Private Sub Form_Load()
Dim x%
    If glbtermopen Then         'Lucy July 5, 2000
        Data1.ConnectionString = glbAdoIHRAUDIT
    Else
        Data1.ConnectionString = glbAdoIHRDB
    End If

    If glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "Select * from Term_HREDSEM"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " AND NOT (ES_CTYPE IS NULL) AND NOT (ES_CRSCODE IS NULL) AND NOT (ES_DATCOMP IS NULL) "
        If Not glbOracle Then SQLQ = SQLQ & " AND LEN(ES_CTYPE)>0 AND LEN(ES_CRSCODE)>0 "
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC,ES_CRSCODE ASC,ES_DATCOMP DESC, ES_EMPNBR"
    Else
        SQLQ = "Select * from HREDSEM"
        SQLQ = SQLQ & " where ES_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND NOT (ES_CTYPE IS NULL) AND NOT (ES_CRSCODE IS NULL) AND NOT (ES_DATCOMP IS NULL) "
        If Not glbOracle Then SQLQ = SQLQ & " AND LEN(ES_CTYPE)>0 AND LEN(ES_CRSCODE)>0 "
        SQLQ = SQLQ & " ORDER BY ES_CTYPE ASC,ES_CRSCODE ASC,ES_DATCOMP DESC, ES_EMPNBR"
    End If

    Data1.RecordSource = SQLQ
    Data1.Refresh

    Call INI_Controls(Me)

    For x% = 0 To 3
        Call setCaption(lbltitle(x%))
    Next
    For x% = 0 To 3
        vbxTrueGrid.Columns(x%).Caption = lStr((vbxTrueGrid.Columns(x%).Caption))
    Next
        
End Sub

Private Sub SaveDataInArray()
    glbCrsCodeStrArr(1) = clpCode(1).Text 'Course Type
    glbCrsCodeStrArr(2) = clpCode(2).Text 'Course Code
    glbCrsCodeStrArr(3) = dlpDatComp.Text '
    glbCrsCodeStrArr(17) = "*" 'Flag
End Sub

Private Sub vbxTrueGrid_DblClick()

    Call SaveDataInArray
    Unload Me

End Sub
