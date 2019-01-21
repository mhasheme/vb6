VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmAttendMatrix 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Attendance Code Matrix"
   ClientHeight    =   7560
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   13530
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
   ScaleHeight     =   7560
   ScaleWidth      =   13530
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPT 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   8280
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4560
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtNotes 
      Appearance      =   0  'Flat
      DataField       =   "AM_NOTES"
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
      Left            =   4560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkNotes 
      Alignment       =   1  'Right Justify
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   9
      Top             =   6405
      Width           =   2115
   End
   Begin VB.TextBox txtFullPartial 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "AM_FULL_PARTIAL"
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
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.ComboBox cmbFullPartial 
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
      Left            =   2250
      TabIndex        =   2
      Tag             =   "11-Choose if Full or Partial Day"
      Top             =   3960
      Width           =   2355
   End
   Begin VB.TextBox txtIncident 
      Appearance      =   0  'Flat
      DataField       =   "AM_INCID"
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
      Left            =   4560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4425
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkRegHrs 
      Alignment       =   1  'Right Justify
      Caption         =   "Regular Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   5
      Top             =   4770
      Width           =   2115
   End
   Begin VB.TextBox txtRegHrs 
      Appearance      =   0  'Flat
      DataField       =   "AM_REG_HRS"
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
      Left            =   4560
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkVacHrs 
      Alignment       =   1  'Right Justify
      Caption         =   "Vacation Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   6
      Top             =   5160
      Width           =   2115
   End
   Begin VB.TextBox txtVacHrs 
      Appearance      =   0  'Flat
      DataField       =   "AM_VAC_HRS"
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
      Left            =   4560
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5175
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkExtraHrs 
      Alignment       =   1  'Right Justify
      Caption         =   "Extra Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   7
      Tag             =   " "
      Top             =   5565
      Width           =   2115
   End
   Begin VB.TextBox txtExtraHrs 
      Appearance      =   0  'Flat
      DataField       =   "AM_EXTRA_HRS"
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
      Left            =   4560
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5535
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkAbsentHrs 
      Alignment       =   1  'Right Justify
      Caption         =   "Absent Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   8
      Top             =   5955
      Width           =   2115
   End
   Begin VB.TextBox txtAbsentHrs 
      Appearance      =   0  'Flat
      DataField       =   "AM_ABSENT_HRS"
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
      Left            =   4560
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5910
      Visible         =   0   'False
      Width           =   495
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "AM_REASON"
      Height          =   285
      Index           =   0
      Left            =   1935
      TabIndex        =   0
      Tag             =   "01-Enter Code for Attendance Reason"
      Top             =   3165
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ADRE"
      MaxLength       =   7
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   6000
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AM_LUSER"
      Enabled         =   0   'False
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
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtCodeType 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "AM_CODE_TYPE"
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
      Left            =   4650
      MaxLength       =   15
      TabIndex        =   17
      Top             =   3540
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AM_LDATE"
      Enabled         =   0   'False
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
      Index           =   0
      Left            =   6600
      MaxLength       =   12
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AM_LTIME"
      Enabled         =   0   'False
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
      Left            =   7350
      MaxLength       =   8
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6240
      Top             =   6120
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
      GridSource      =   "vbxTrueGrid"
   End
   Begin VB.ComboBox cmbCodeType 
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
      Left            =   2250
      TabIndex        =   1
      Tag             =   "11-Choose Type of Attendance Reason Code"
      Top             =   3545
      Width           =   2355
   End
   Begin VB.CheckBox chkIncident 
      Alignment       =   1  'Right Justify
      Caption         =   "Incident"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   330
      TabIndex        =   4
      Top             =   4365
      Width           =   2115
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feAttMatrix.frx":0000
      Height          =   2895
      Left            =   0
      OleObjectBlob   =   "feAttMatrix.frx":0014
      TabIndex        =   11
      Top             =   120
      Width           =   11535
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   7920
      TabIndex        =   3
      Tag             =   "00-Section - Code"
      Top             =   3960
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      DataField       =   "AM_PT"
      Height          =   285
      Left            =   1935
      TabIndex        =   10
      Tag             =   "EDPT-Category"
      Top             =   6960
      Visible         =   0   'False
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
      Left            =   7440
      TabIndex        =   30
      Tag             =   "00-Section - Code"
      Top             =   5400
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblSection2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section Filter"
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
      Left            =   6000
      TabIndex        =   31
      Top             =   5400
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   28
      Top             =   7005
      Visible         =   0   'False
      Width           =   765
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
      Left            =   6480
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Full or Partial Day"
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
      Left            =   330
      TabIndex        =   24
      Top             =   4020
      Width           =   1230
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   16
      Top             =   3205
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   15
      Top             =   3578
      Width           =   930
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "AM_COMPNO"
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
      Left            =   6720
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmAttendMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum

Private Function chkAttMatrix()
Dim Msg As String
Dim x%, xchk
Dim rsAttMat As New ADODB.Recordset
Dim SQLQ As String
Dim xCatCount As Integer
Dim xCatSource As String
Dim xCatSearch() As String
Dim I As Integer

chkAttMatrix = False

If Len(clpCode(0).Text) < 1 Then
    MsgBox "Reason Code must be entered"
    clpCode(0).SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
    MsgBox "Invalid Reason Code"
    clpCode(0).SetFocus
    Exit Function
End If

'Ticket #24586 - WDGPHU
If glbCompSerial = "S/N - 2411W" Then
    If Len(clpPT.Text) < 1 Then
        MsgBox lStr("Category") & " must be entered"
        clpPT.SetFocus
        Exit Function
    End If
    
    If Not clpPT.ListChecker Then
        'MsgBox "Invalid " & lStr("Category")
        'clpPT.SetFocus
        Exit Function
    End If
        
    'Check if the duplicate Reason Code already exists
    SQLQ = "SELECT * FROM HRATT_MATRIX"
    SQLQ = SQLQ$ & " WHERE AM_REASON = '" & clpCode(0).Text & "'"
    If Not fglbNew Then
        SQLQ = SQLQ & " AND AM_ID <> " & Data1.Recordset!AM_ID
    End If
    SQLQ = SQLQ & " ORDER BY AM_REASON "
    rsAttMat.Open SQLQ$, gdbAdoIhr001, adOpenStatic
    If rsAttMat.EOF Then
        rsAttMat.Close
        Set rsAttMat = Nothing
    Else
        MsgBox "The Attendance Code Matrix for this Reason already exists."
        rsAttMat.Close
        Set rsAttMat = Nothing
        clpCode(0).SetFocus
        Exit Function
    End If
    
    'Check if the Category Code does not exists multiple times in the Category List
    'InstrCount = Len(Replace(Source, Search, Search & "*")) - Len(Source)
    xCatCount = 0
    xCatSource = ""
    xCatSource = "'" & Replace(clpPT.Text, ",", "','") & "'"
    xCatSearch = Split(xCatSource, ",")
    For I = 0 To UBound(xCatSearch)
        xCatCount = Len(Replace(xCatSource, xCatSearch(I), xCatSearch(I) & "*")) - Len(xCatSource)
        
        If xCatCount > 1 Then
            MsgBox lStr("Category") & " " & xCatSearch(I) & " is defined more than once in the list."
            Exit Function
            Exit For
        End If
    Next I
Else
    If cmbCodeType.ListIndex = -1 Then
        MsgBox "Please select Code Type"
        cmbCodeType.SetFocus
        Exit Function
    End If
    
    'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
    If glbWFC Or glbCompSerial = "S/N - 2335W" Then  'Ticket #24112 Franks 07/30/2013  for Mitchell Plastics
        If Len(clpCode(1).Text) < 1 Then
            MsgBox lStr("Section") & " Code must be entered"
            clpCode(1).SetFocus
            Exit Function
        End If
        If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
            MsgBox "Invalid " & lStr("Section") & " Code"
            clpCode(1).SetFocus
            Exit Function
        End If
    End If
    If glbWFC Then 'Ticket #28373 Franks 04/28/2016
        'Ticket #28664 Franks 05/30/2016 "   Category not mandatory.
        'If Len(clpPT.Text) < 1 Then
        '    MsgBox lStr("Category") & " must be entered"
        '    clpPT.SetFocus
        '    Exit Function
        'End If
        
        If Not clpPT.ListChecker Then
            'MsgBox "Invalid " & lStr("Category")
            'clpPT.SetFocus
            Exit Function
        End If
    End If
End If

chkAttMatrix = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fglbNew = False

If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If

rsDATA.CancelUpdate

Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRATT_MATRIX", "Cancel")
Call RollBack '09June99 js

End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRATT_MATRIX", "Delete")
Call RollBack '09June99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

'Ticket #24586 - WDGPHU
If glbCompSerial <> "S/N - 2411W" Then
    clpCode(0).SetFocus
End If

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRATT_MATRIX", "Modify")
Call RollBack '09June99 js

End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)

rsDATA.AddNew

lblCNum.Caption = "001"

'Ticket #24586 - WDGPHU
If glbCompSerial <> "S/N - 2411W" Then
    cmbCodeType.ListIndex = -1
    cmbFullPartial.ListIndex = -1
End If

fglbNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
clpCode(0).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRATT_MATRIX", "Add")
Call RollBack '09June99 js

End Sub

Sub cmdOK_Click()
Dim x%
Dim bmk As Variant

On Error GoTo cmdOK_Err

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkAttMatrix() Then Exit Sub

'If glbCompSerial = "S/N - 2411W" Then   'Ticket #24586 - WDGPHU
'    txtPT.Text = "'" & Replace(clpPT.Text, ",", "','") & "'"
'    clpPT.DataField = ""
'End If

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

fglbNew = False

Call Display_Value

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRATT_MATRIX", "Update")
Call RollBack '09June99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Attendance Code Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

RHeading = "Attendance Code Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmbCodeType_AddItems()
    'Ticket #24586 - WDGPHU
    If glbCompSerial <> "S/N - 2411W" Then
    
        'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
        If glbWFC Or glbCompSerial = "S/N - 2335W" Then  'Ticket #24112 Franks 07/30/2013  for Mitchell Plastics
            cmbCodeType.Clear
            cmbCodeType.AddItem "Paid"
            cmbCodeType.AddItem "Unpaid"
        ElseIf glbCompSerial = "S/N - 2370W" Then   'Chapmans
            cmbCodeType.Clear
            cmbCodeType.AddItem "Bereavement"
            cmbCodeType.AddItem "Emergency Leave"
            cmbCodeType.AddItem "Excused"
            cmbCodeType.AddItem "Family Sick"
            cmbCodeType.AddItem "Late"
            cmbCodeType.AddItem "Left Early"
            cmbCodeType.AddItem "Other or N/A"
            cmbCodeType.AddItem "Sick"
            cmbCodeType.AddItem "Weather"
        Else    'Ticket #25922 - OHRS Reporting for CHC
            cmbCodeType.Clear
            cmbCodeType.AddItem "Worked Hours"
            cmbCodeType.AddItem "Benefit Hours"
        End If
        
        cmbFullPartial.Clear
        cmbFullPartial.AddItem ""
        cmbFullPartial.AddItem "Full Day"
        cmbFullPartial.AddItem "Partial Day"
    End If
End Sub

Private Sub chkAbsentHrs_Click()
    If chkAbsentHrs.Value = 1 Then
        txtAbsentHrs.Text = "1"
    Else
        txtAbsentHrs.Text = "0"
    End If
End Sub

Private Sub chkExtraHrs_Click()
    If chkExtraHrs.Value = 1 Then
        txtExtraHrs.Text = "1"
    Else
        txtExtraHrs.Text = "0"
    End If
End Sub

Private Sub chkIncident_Click()
    If chkIncident.Value = 1 Then
        txtIncident.Text = "1"
    Else
        txtIncident.Text = "0"
    End If
End Sub

Private Sub chkNotes_Click()
    If chkNotes.Value = 1 Then
        txtNotes.Text = "1"
    Else
        txtNotes.Text = "0"
    End If
End Sub

Private Sub chkRegHrs_Click()
    If chkRegHrs.Value = 1 Then
        txtRegHrs.Text = "1"
    Else
        txtRegHrs.Text = "0"
    End If
End Sub

Private Sub chkVacHrs_Click()
    If chkVacHrs.Value = 1 Then
        txtVacHrs.Text = "1"
    Else
        txtVacHrs.Text = "0"
    End If
End Sub

Private Sub clpCode_Change(Index As Integer)
If glbWFC And Index = 2 Then 'Ticket #27298 Franks 07/13/2015
Dim SQLQ
    SQLQ = "SELECT * FROM HRATT_MATRIX "
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & "WHERE AM_SECTION = '" & clpCode(2).Text & "' "
    End If
    SQLQ = SQLQ & " ORDER BY AM_SECTION, AM_REASON,AM_CODE_TYPE "
    Data1.RecordSource = SQLQ
    Data1.Refresh
End If

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpPT_Change()
    'txtPT.Text = "'" & Replace(clpPT.Text, ",", "','") & "'"
End Sub

Private Sub clpPT_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbCodeType_Click()
'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
If glbWFC Or glbCompSerial = "S/N - 2335W" Then  'Ticket #24112 Franks 07/30/2013  for Mitchell Plastics
    Select Case cmbCodeType.ListIndex
        Case 0: txtCodeType.Text = "Paid"
        Case 1: txtCodeType.Text = "Unpaid"
    End Select
ElseIf glbCompSerial = "S/N - 2370W" Then  'Chapmans
    Select Case cmbCodeType.ListIndex
        Case 0: txtCodeType.Text = "BRV"
        Case 1: txtCodeType.Text = "EML"
        Case 2: txtCodeType.Text = "EXC"
        Case 3: txtCodeType.Text = "FSC"
        Case 4: txtCodeType.Text = "LAT"
        Case 5: txtCodeType.Text = "LFE"
        Case 6: txtCodeType.Text = "OTH"
        Case 7: txtCodeType.Text = "SIC"
        Case 8: txtCodeType.Text = "WTH"
    End Select
Else    'Ticket #25922 - OHRS Reporting for CHC
    Select Case cmbCodeType.ListIndex
        Case 0: txtCodeType.Text = "Worked Hours"
        Case 1: txtCodeType.Text = "Benefit Hours"
    End Select
End If
End Sub

Private Sub cmbCodeType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbFullPartial_Click()
    Select Case cmbFullPartial.ListIndex
        Case 0: txtFullPartial.Text = ""
        Case 1: txtFullPartial.Text = "F"
        Case 2: txtFullPartial.Text = "P"
    End Select
End Sub

Private Sub cmbFullPartial_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "ATTENDANCE CODE MATRIX", "SELECT")

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%, SQLQ

'Me.Show

glbOnTop = "FRMATTENDMATRIX"

Screen.MousePointer = HOURGLASS

'Ticket #24124 Franks 07/24/2013 - begin
If glbWFC Or glbCompSerial = "S/N - 2335W" Then  'Ticket #24112 Franks 07/30/2013  for Mitchell Plastics
    Call WFCScreenSetup
ElseIf glbCompSerial = "S/N - 2411W" Then   'Ticket #24586 - WDGPHU
    Call WDGPHUScreenSetup
ElseIf glbCompSerial = "S/N - 2370W" Then  'Chapmans
    vbxTrueGrid.Columns(9).Visible = False
ElseIf glbCompSerial = "S/N - 2485W" Then   'Ticket #29617 - Mississaugas of Scugog Island First Nation
    Call MScugogFirstNationScreenSetup
Else
    Call RestOfClientsSetup
End If

Call cmbCodeType_AddItems
'Ticket #24124 Franks 07/24/2013 - end

Data1.ConnectionString = glbAdoIHRDB

'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
If glbWFC Or glbCompSerial = "S/N - 2335W" Then  'Ticket #24112 Franks 07/30/2013  for Mitchell Plastics
    Data1.RecordSource = "SELECT * FROM HRATT_MATRIX ORDER BY AM_SECTION, AM_REASON,AM_CODE_TYPE "
Else
    Data1.RecordSource = "SELECT * FROM HRATT_MATRIX ORDER BY AM_REASON,AM_CODE_TYPE "
End If
Data1.Refresh

Screen.MousePointer = DEFAULT

'Call Display_Value

Call ST_UPD_MODE(False)

'vbxTrueGrid.Columns(0).Caption = lStr("Section")
'lblTitle(0).Caption = lStr(lblTitle(0).Caption)
'lblTitle(1).Caption = lStr(lblTitle(1).Caption)
'lblTitle(2).Caption = lStr(lblTitle(2).Caption)
Call INI_Controls(Me)

Screen.MousePointer = DEFAULT                           '

End Sub

Private Sub Form_LostFocus()

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

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
Dim I As Integer
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

fUPMode = TF
'vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF              '
'cmdCancel.Enabled = TF          '
'cmdClose.Enabled = FT           '
'cmdModify.Enabled = FT          '
'cmdNew.Enabled = FT             '
'cmdDelete.Enabled = FT          '
'cmdPrint.Enabled = FT           '
clpCode(0).Enabled = TF
cmbCodeType.Enabled = TF            '
cmbFullPartial.Enabled = TF                    '
chkAbsentHrs.Enabled = TF      '
chkExtraHrs.Enabled = TF      '
chkIncident.Enabled = TF      '
chkRegHrs.Enabled = TF      '
chkVacHrs.Enabled = TF
chkNotes.Enabled = TF
clpCode(1).Enabled = TF 'Ticket #24124 Franks 07/24/2013

'Ticket #24586 - WDGPHU
If glbCompSerial = "S/N - 2411W" Or glbWFC Then
    clpPT.Enabled = TF
End If
End Sub

Private Sub txtAbsentHrs_Change()
    If txtAbsentHrs = "-1" Or txtAbsentHrs = "1" Then
        chkAbsentHrs = 1
    Else
        chkAbsentHrs = 0
    End If
End Sub

Private Sub txtCodeType_Change()
'If glbWFC Then 'Ticket #24124 Franks 07/24/2013
If glbWFC Or glbCompSerial = "S/N - 2335W" Then  'Ticket #24112 Franks 07/30/2013  for Mitchell Plastics
    If txtCodeType = "Paid" Then
        cmbCodeType.ListIndex = 0
    End If
    
    If txtCodeType = "Unpaid" Then
        cmbCodeType.ListIndex = 1
    End If
ElseIf glbCompSerial <> "S/N - 2370W" And glbCompSerial <> "S/N - 2411W" Then  'Not Chapmans and not WDGPHU
    'Ticket #25922 - OHRS Reporting for CHC
    If txtCodeType = "Worked Hours" Then
        cmbCodeType.ListIndex = 0
    End If
    
    If txtCodeType = "Benefit Hours" Then
        cmbCodeType.ListIndex = 1
    End If
    
ElseIf glbCompSerial <> "S/N - 2411W" Then  'Ticket #24586 - WDGPHU
    If txtCodeType = "BRV" Then
        cmbCodeType.ListIndex = 0
    End If
    
    If txtCodeType = "EML" Then
        cmbCodeType.ListIndex = 1
    End If
    
    If txtCodeType = "EXC" Then
        cmbCodeType.ListIndex = 2
    End If
    
    If txtCodeType = "FSC" Then
        cmbCodeType.ListIndex = 3
    End If
    
    If txtCodeType = "LAT" Then
        cmbCodeType.ListIndex = 4
    End If
    
    If txtCodeType = "LFE" Then
        cmbCodeType.ListIndex = 5
    End If
    
    If txtCodeType = "OTH" Then
        cmbCodeType.ListIndex = 6
    End If
    
    If txtCodeType = "SIC" Then
        cmbCodeType.ListIndex = 7
    End If
    
    If txtCodeType = "WTH" Then
        cmbCodeType.ListIndex = 8
    End If
End If

End Sub

Private Sub txtExtraHrs_Change()
    If txtExtraHrs = "-1" Or txtExtraHrs = "1" Then
        chkExtraHrs = 1
    Else
        chkExtraHrs = 0
    End If
End Sub

Private Sub txtFullPartial_Change()
If glbCompSerial <> "S/N - 2411W" Then  'Ticket #24586 - WDGPHU
    If txtFullPartial = "F" Then
        cmbFullPartial.ListIndex = 1
    End If
    
    If txtFullPartial = "P" Then
        cmbFullPartial.ListIndex = 2
    End If
    
    If txtFullPartial = "" Then
        cmbFullPartial.ListIndex = 0
    End If
End If
End Sub

Private Sub txtIncident_Change()
    If txtIncident = "-1" Or txtIncident = "1" Then
        chkIncident = 1
    Else
        chkIncident = 0
    End If
End Sub

Private Sub txtNotes_Change()
    If txtNotes = "-1" Or txtNotes = "1" Then
        chkNotes = 1
    Else
        chkNotes = 0
    End If
End Sub

Private Sub txtPT_Change()
    'clpPT.Text = Replace(txtPT.Text, "'", "")
End Sub

Private Sub txtRegHrs_Change()
    If txtRegHrs = "-1" Or txtRegHrs = "1" Then
        chkRegHrs = 1
    Else
        chkRegHrs = 0
    End If
End Sub

Private Sub txtVacHrs_Change()
    If txtVacHrs = "-1" Or txtVacHrs = "1" Then
        chkVacHrs = 1
    Else
        chkVacHrs = 0
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
    
    SQLQ = "SELECT * FROM HRATT_MATRIX "
    If glbWFC Then 'Ticket #27298 Franks 07/13/2015
        If Len(clpCode(2).Text) > 0 Then
            SQLQ = SQLQ & "WHERE AM_SECTION = '" & clpCode(2).Text & "' "
        End If
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If

If glbCompSerial <> "S/N - 2411W" Then  'Ticket #24586 - WDGPHU

    If glbCompSerial <> "S/N - 2370W" Then  'Not Chapmans and not WDGPHU
        'Ticket #25922 - OHRS Reporting for CHC
        If txtCodeType = "Worked Hours" Then
            cmbCodeType.ListIndex = 0
        End If
        
        If txtCodeType = "Benefit Hours" Then
            cmbCodeType.ListIndex = 1
        End If
    Else
        If txtCodeType = "BRV" Then
            cmbCodeType.ListIndex = 0
        End If
        
        If txtCodeType = "EML" Then
            cmbCodeType.ListIndex = 1
        End If
        
        If txtCodeType = "EXC" Then
            cmbCodeType.ListIndex = 2
        End If
        
        If txtCodeType = "FSC" Then
            cmbCodeType.ListIndex = 3
        End If
        
        If txtCodeType = "LAT" Then
            cmbCodeType.ListIndex = 4
        End If
        
        If txtCodeType = "LFE" Then
            cmbCodeType.ListIndex = 5
        End If
        
        If txtCodeType = "OTH" Then
            cmbCodeType.ListIndex = 6
        End If
        
        If txtCodeType = "SIC" Then
            cmbCodeType.ListIndex = 7
        End If
        
        If txtCodeType = "WTH" Then
            cmbCodeType.ListIndex = 8
        End If
    End If
End If

Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRATT_MATRIX", "Select")
Call RollBack '09June99 js

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

Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRATT_MATRIX where AM_ID= " & Data1.Recordset!AM_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
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
UpdateRight = gSec_Upd_AttendCode_Matrix
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
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

End Sub

Private Sub WFCScreenSetup() 'Ticket #24124 Franks 07/24/2013
    'hide unused fields - begin
    lblTitle(2).Visible = False
    cmbFullPartial.Visible = False
    chkIncident.Visible = False
    chkRegHrs.Visible = False
    chkVacHrs.Visible = False
    chkExtraHrs.Visible = False
    
    'Ticket #27298 Franks 07/13/2015 - begin
    'chkAbsentHrs.Visible = False
    chkAbsentHrs.Caption = "Using Time Banks"
    chkAbsentHrs.Left = 300
    chkAbsentHrs.Top = cmbFullPartial.Top ' chkIncident.Top
    'vbxTrueGrid.Columns(10).Visible = False
    'Ticket #27298 Franks 07/13/2015 - end
    
    chkNotes.Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(3).Visible = False
    vbxTrueGrid.Columns(4).Visible = False
    vbxTrueGrid.Columns(5).Visible = False
    vbxTrueGrid.Columns(6).Visible = False
    'vbxTrueGrid.Columns(7).Visible = False 'Ticket #27298 Franks 07/13/2015
    vbxTrueGrid.Columns(7).Caption = "Using Time Banks"
    
    vbxTrueGrid.Columns(8).Visible = False
    'hide unused fields - end
    'Section for WFC
    lblSection.Left = lblTitle(2).Left
    lblSection.Top = chkIncident.Top ' lblTitle(2).Top
    clpCode(1).Left = clpCode(0).Left
    clpCode(1).Top = chkIncident.Top ' cmbFullPartial.Top
    lblSection.FontBold = True
    lblSection.Caption = lStr("Section")
    vbxTrueGrid.Columns(9).Caption = lStr("Section")
    lblSection.Visible = True
    clpCode(1).Visible = True
    clpCode(1).DataField = "AM_SECTION"
    
    'Section Filter
    lblSection2.Caption = lStr("Section") & " Filter"
    lblSection2.Top = lblSection.Top
    clpCode(2).Top = clpCode(1).Top
    lblSection2.Visible = True
    clpCode(2).Visible = True

    'Category
    lblPT.Left = lblTitle(0).Left
    lblPT.Top = chkRegHrs.Top
    clpPT.Left = clpCode(0).Left
    clpPT.Top = chkRegHrs.Top
    lblPT.Caption = lStr("Category")
    lblPT.FontBold = False
    vbxTrueGrid.Columns(10).Caption = lStr("Category")
    lblPT.Visible = True
    clpPT.Visible = True
    
End Sub

Private Sub WDGPHUScreenSetup() 'Ticket #24586 - WDGPHU
    'Hide unused fields - begin
    lblTitle(0).Visible = False
    cmbCodeType.Visible = False
    lblTitle(2).Visible = False
    cmbFullPartial.Visible = False
    chkIncident.Visible = False
    chkRegHrs.Visible = False
    chkVacHrs.Visible = False
    chkExtraHrs.Visible = False
    chkAbsentHrs.Visible = False
    chkNotes.Visible = False
    vbxTrueGrid.Columns(1).Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(3).Visible = False
    vbxTrueGrid.Columns(4).Visible = False
    vbxTrueGrid.Columns(5).Visible = False
    vbxTrueGrid.Columns(6).Visible = False
    vbxTrueGrid.Columns(7).Visible = False
    vbxTrueGrid.Columns(8).Visible = False
    vbxTrueGrid.Columns(9).Visible = False
    'Hide unused fields - end
    
    'Category for WDGPHU
    lblPT.Left = lblTitle(0).Left
    lblPT.Top = lblTitle(0).Top
    clpPT.Left = clpCode(0).Left
    clpPT.Top = cmbCodeType.Top
    lblPT.FontBold = True
    lblPT.Caption = lStr("Category")
    vbxTrueGrid.Columns(10).Caption = lStr("Category")
    lblPT.Visible = True
    clpPT.Visible = True
    'clpPT.DataField = "AM_PT"
    'txtPT.DataField = "AM_PT"
End Sub

Private Sub MScugogFirstNationScreenSetup()
    Dim xSwitch As Integer
    
    'Show fields and reposition it
    lblPT.Left = lblTitle(0).Left
    lblPT.Top = lblTitle(0).Top
    clpPT.Left = clpCode(0).Left
    clpPT.Top = cmbCodeType.Top
    lblPT.FontBold = True
    lblPT.Caption = lStr("Category")
    vbxTrueGrid.Columns(10).Caption = lStr("Category")
    lblPT.Visible = True
    clpPT.Visible = True

    lblTitle(0).Left = lblTitle(2).Left
    lblTitle(0).Top = lblTitle(2).Top
    cmbCodeType.Left = cmbFullPartial.Left
    cmbCodeType.Top = cmbFullPartial.Top

    'Hide unused fields - begin
    lblTitle(2).Visible = False
    cmbFullPartial.Visible = False
    chkNotes.Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(8).Visible = False
    vbxTrueGrid.Columns(9).Visible = False
    'Hide unused fields - end
    
    'Relabel / Rearrange fields
    chkIncident.Caption = "IW+ (Inclement Weather)"  'AM_INCID
    chkRegHrs.Caption = "PT+ (Paid Time Off)"       'AM_REG_HRS
    chkVacHrs.Caption = "VA+ (Vacation)"             'AM_VAC_HRS
    chkExtraHrs.Caption = "PD+ (Personal Day)"       'AM_EXTRA_HRS
    chkAbsentHrs.Caption = "SK+ (Sick)"              'AM_ABSENT_HRS
    
    xSwitch = chkAbsentHrs.Top
    chkAbsentHrs.Top = chkExtraHrs.Top
    chkExtraHrs.Top = chkVacHrs.Top
    chkVacHrs.Top = xSwitch
    xSwitch = chkExtraHrs.Top
    chkExtraHrs.Top = chkRegHrs.Top
    chkRegHrs.Top = xSwitch
    
    'chkRegHrs.Top = chkVacHrs.Top                   '4
    'chkVacHrs.Top = chkAbsentHrs.Top                '1
    'chkAbsentHrs.Top = chkExtraHrs.Top              '2
    'chkExtraHrs.Top = chkRegHrs.Top                 '3
    'chkIncident.Top = chkAbsentHrs.Top             '5
    
    vbxTrueGrid.Columns(3).Caption = "IW+"
    vbxTrueGrid.Columns(4).Caption = "PT+"
    vbxTrueGrid.Columns(5).Caption = "VA+"
    vbxTrueGrid.Columns(6).Caption = "PD+"
    vbxTrueGrid.Columns(7).Caption = "SK+"

    vbxTrueGrid.Columns(1).Order = 2
    vbxTrueGrid.Columns(10).Order = 1
    vbxTrueGrid.Columns(3).Order = 5
    vbxTrueGrid.Columns(4).Order = 6
    vbxTrueGrid.Columns(5).Order = 8
    vbxTrueGrid.Columns(6).Order = 5
    vbxTrueGrid.Columns(7).Order = 7
    
    'Tab Order setup
    vbxTrueGrid.TabIndex = 0
    clpCode(0).TabIndex = 1
    clpPT.TabIndex = 2
    cmbCodeType.TabIndex = 3
    chkIncident.TabIndex = 4
    chkExtraHrs.TabIndex = 5
    chkRegHrs.TabIndex = 6
    chkAbsentHrs.TabIndex = 7
    chkVacHrs.TabIndex = 8
End Sub

Private Sub RestOfClientsSetup()
    'hide unused fields - begin
    lblTitle(2).Visible = False
    cmbFullPartial.Visible = False
    chkIncident.Visible = False
    chkRegHrs.Visible = False
    chkVacHrs.Visible = False
    chkExtraHrs.Visible = False
    chkAbsentHrs.Visible = False
    chkNotes.Visible = False
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(3).Visible = False
    vbxTrueGrid.Columns(4).Visible = False
    vbxTrueGrid.Columns(5).Visible = False
    vbxTrueGrid.Columns(6).Visible = False
    vbxTrueGrid.Columns(7).Visible = False
    vbxTrueGrid.Columns(8).Visible = False
    vbxTrueGrid.Columns(9).Visible = False
    vbxTrueGrid.Columns(10).Visible = False
End Sub
