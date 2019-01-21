VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmPAYROLL 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Payroll Matrix"
   ClientHeight    =   6885
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   10560
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
   ScaleHeight     =   6885
   ScaleWidth      =   10560
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbConvert2 
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
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Tag             =   "11-Choose Convert 2 of Payroll Matrix"
      Top             =   5520
      Visible         =   0   'False
      Width           =   4275
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   15
      Tag             =   "00-Administered By"
      Top             =   4440
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin VB.ComboBox cmbConvert3 
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
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Tag             =   "11-Choose Convert 4 of Payroll Matrix"
      Top             =   4800
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.ComboBox cmbConvert4 
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
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "11-Choose Convert 4 of Payroll Matrix"
      Top             =   5160
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox txtNumConvert 
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
      Index           =   2
      Left            =   3090
      MaxLength       =   20
      TabIndex        =   20
      Tag             =   "00-Enter Conversion Number #2"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox txtNumConvert 
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
      Index           =   1
      Left            =   3090
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "00-Enter Conversion Number #1"
      Top             =   5520
      Width           =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "M_CODE"
      Height          =   285
      Index           =   0
      Left            =   2775
      TabIndex        =   11
      Tag             =   "01-Enter Code for Attendance"
      Top             =   3540
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ADRE"
      MaxLength       =   7
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fepayrol.frx":0000
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "fepayrol.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   10215
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "M_CODEDEPT"
      Height          =   285
      Left            =   2775
      TabIndex        =   12
      Tag             =   "00-Specific Department Desired"
      Top             =   3870
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6840
      Top             =   6120
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
      DataField       =   "M_LUSER"
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
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtTransfer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "M_DEFTYPE"
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
      Left            =   7230
      MaxLength       =   4
      TabIndex        =   23
      Top             =   3900
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtType 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "M_TYPE"
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
      Left            =   6210
      MaxLength       =   4
      TabIndex        =   22
      Top             =   3900
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtConvert 
      Appearance      =   0  'Flat
      DataField       =   "M_CONVERT4"
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
      Index           =   3
      Left            =   3090
      MaxLength       =   10
      TabIndex        =   17
      Tag             =   "00-Enter Conversion #4"
      Top             =   5190
      Width           =   1215
   End
   Begin VB.TextBox txtConvert 
      Appearance      =   0  'Flat
      DataField       =   "M_CONVERT3"
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
      Index           =   2
      Left            =   3090
      MaxLength       =   10
      TabIndex        =   16
      Tag             =   "00-Enter Conversion #3"
      Top             =   4860
      Width           =   1215
   End
   Begin VB.TextBox txtConvert 
      Appearance      =   0  'Flat
      DataField       =   "M_CONVERT2"
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
      Index           =   1
      Left            =   3090
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "00-Enter Conversion #2"
      Top             =   4530
      Width           =   1215
   End
   Begin VB.TextBox txtConvert 
      Appearance      =   0  'Flat
      DataField       =   "M_CONVERT1"
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
      Left            =   3090
      MaxLength       =   10
      TabIndex        =   13
      Tag             =   "00-Enter Conversion #1"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox cmbType 
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
      Left            =   3090
      TabIndex        =   10
      Tag             =   "11-Choose Type of Payroll Matrix"
      Top             =   3150
      Width           =   1515
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "M_LDATE"
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
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3465
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "M_LTIME"
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
      Left            =   7590
      MaxLength       =   8
      TabIndex        =   2
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "M_SECTION"
      Height          =   285
      Index           =   1
      Left            =   2775
      TabIndex        =   21
      Top             =   6240
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblNote2 
      Caption         =   "Note2"
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
      Left            =   4560
      TabIndex        =   29
      Top             =   4560
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert Number 2"
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
      Left            =   330
      TabIndex        =   28
      Top             =   5880
      Width           =   2400
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert Number 1"
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
      Left            =   330
      TabIndex        =   27
      Top             =   5520
      Width           =   2520
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
      Left            =   330
      TabIndex        =   26
      Top             =   6240
      Width           =   2085
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   330
      TabIndex        =   25
      Top             =   3870
      Width           =   990
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert 4"
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
      Left            =   330
      TabIndex        =   9
      Top             =   5220
      Width           =   2160
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert 3"
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
      Index           =   4
      Left            =   330
      TabIndex        =   8
      Top             =   4890
      Width           =   2760
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert 2"
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
      Left            =   330
      TabIndex        =   7
      Top             =   4560
      Width           =   2760
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Convert 1"
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
      TabIndex        =   6
      Top             =   4230
      Width           =   2760
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "INFO:HR Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   5
      Top             =   3540
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   4
      Top             =   3210
      Width           =   435
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "M_COMPNO"
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
      Left            =   9600
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmPAYROLL"
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
Dim NewAccpacMatrix As Boolean 'Ticket #24267 Franks 01/14/2014
Dim NewPayWebMatrix As Boolean 'Let's Talk Science Ticket #27072 10/14/2015

Private Function chkEPayroll()
Dim Msg As String
Dim x%, xchk

chkEPayroll = False
If cmbType <> "DEPT" Then
    If Len(clpCode(0).Text) < 1 Then
        MsgBox "info:HR Code must be entered"
         clpCode(0).SetFocus
        Exit Function
    End If
    
    If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox "Code entered must be known"
         clpCode(0).SetFocus
        Exit Function
    End If
    If glbWFC Then
        'Ticket #16392 US Payroll not need Plant Code
        'If Len(clpCode(1).Text) < 1 Then
        '    MsgBox lStr("Section Code must be entered")
        '    clpCode(1).SetFocus
        '    Exit Function
        'End If
    End If
    If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
        MsgBox "Code entered must be known"
        clpCode(1).SetFocus
        Exit Function
    End If
Else
    If Len(clpDept.Text) = 0 Then
        MsgBox "Department must be entered"
         clpDept.SetFocus
        Exit Function
    End If
    If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
        MsgBox "If Department Entered - it must be known"
         clpDept.SetFocus
        Exit Function
    End If

End If
If cmbType = "PAYT" Then
    If glbCompSerial = "S/N - 2382W" Then 'Namasco
        If Len(txtConvert(1).Text) > 0 Then '"Earnings/Deduction Indicator(E or D)" 'Convert 2
            If Not (txtConvert(1).Text = "E" Or txtConvert(1).Text = "D") Then
                MsgBox "Valid values are 'E' or 'D'"
                txtConvert(1).SetFocus
                Exit Function
            End If
        End If
        If Len(txtConvert(2).Text) > 0 Then 'Company Code 'Convert 3
            If Not (txtConvert(2).Text = "QWF" Or txtConvert(2).Text = "QWG") Then
                MsgBox "Valid values are 'QWF' or 'QWG'"
                txtConvert(2).SetFocus
                Exit Function
            End If
        End If
    End If
End If


'Hemu 05/13/2003 Begin
If clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
    clpDept.SetFocus
    Exit Function
End If
'Hemu 05/13/2003 End

xchk = False

For x% = 0 To 3
    If Len(txtConvert(x%)) > 0 Then xchk = True
Next x%

If xchk = False Then
    MsgBox "At least one Convert Code must be entered"
    txtConvert(0).SetFocus
    Exit Function
End If

If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #20696 Franks 10/14/2011
    If cmbType.Text = "BENE" Then
        If Len(cmbConvert4.Text) = 0 Then
            MsgBox "Convert 4 is required for Type 'BENE'"
            txtConvert(3).SetFocus
            Exit Function
        End If
    End If
End If
    
If glbCompSerial = "S/N - 2439W" Then  'OK Tire  - Ticket #21519 Franks 06/14/2012
    If cmbType.Text = "BENE" Then
        If Len(cmbConvert4.Text) = 0 Then
            MsgBox "Convert 4 is required for Type 'BENE'"
            txtConvert(3).SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2436W" Then  'Family Day  - Ticket #21152 Franks 11/22/2012
    If cmbType.Text = "BENE" Then
        If Len(cmbConvert3.Text) = 0 Then 'Ticket #23779 Franks 05/30/2013
            MsgBox "Convert 3 is required for Type 'BENE'"
            cmbConvert3.SetFocus
            Exit Function
        End If
        If Len(cmbConvert4.Text) = 0 Then
            MsgBox "Convert 4 is required for Type 'BENE'"
            cmbConvert4.SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2437W" Then  'KN&V  - Ticket #21096 Franks 12/18/2012
    If cmbType.Text = "BENE" Then
        If Len(cmbConvert4.Text) = 0 Then
            MsgBox "Convert 4 is required for Type 'BENE'"
            txtConvert(3).SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2457W" Then 'McLeod Law  Ticket #24864 Franks 11/11/2014
    If cmbType.Text = "BENE" Then
        If Len(txtConvert(0).Text) = 0 Then
            MsgBox "Pay Code is required for Type 'BENE'"
            txtConvert(0).SetFocus
            Exit Function
        End If
        If Len(clpCode(2).Text) = 0 Then
            MsgBox lStr("Administered By") & " is required for Type 'BENE'"
            clpCode(2).SetFocus
            Exit Function
        End If
        If Len(cmbConvert4.Text) = 0 Then
            MsgBox "Amount is required for Type 'BENE'"
            txtConvert(3).SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
    If cmbType.Text = "BENE" Then
        If Len(txtConvert(0).Text) = 0 Then
            MsgBox "Pay Code is required for Type 'BENE'"
            txtConvert(0).SetFocus
            Exit Function
        End If
        If Len(clpCode(2).Text) = 0 Then
            MsgBox lStr("Administered By") & " is required for Type 'BENE'"
            clpCode(2).SetFocus
            Exit Function
        End If
        If Len(cmbConvert4.Text) = 0 Then
            MsgBox "Amount is required for Type 'BENE'"
            txtConvert(3).SetFocus
            Exit Function
        End If
    End If
End If

If NewPayWebMatrix Then 'Let's Talk Science Ticket #27072 10/14/2015
    If cmbType.Text = "BENE" Then
        If Len(txtConvert(0).Text) = 0 Then
            MsgBox lblTitle(2).Caption & " is required for Type 'BENE'"
            txtConvert(0).SetFocus
            Exit Function
        End If
        If cmbConvert4.Visible Then
            If Len(cmbConvert4.Text) = 0 Then
                MsgBox lblTitle(5).Caption & " is required for Type 'BENE'"
                cmbConvert4.SetFocus
                Exit Function
            End If
        End If
        If Len(txtNumConvert(1).Text) = 0 Then
            MsgBox lblTitle(6).Caption & " is required for Type 'BENE'" & Chr(10) & Chr(10) & "Note: enter 0 if no Add Tax"
            txtNumConvert(1).SetFocus
            Exit Function
        End If
    End If
End If

chkEPayroll = True

End Function

Private Sub clpCode_Change(Index As Integer)
    If glbCompSerial = "S/N - 2457W" Then 'Ticket #24864 Franks 11/11/2014
        If Index = 2 Then
            txtConvert(1).Text = clpCode(2).Text
        End If
    End If
    
    If glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
        If Index = 2 Then
            txtConvert(1).Text = clpCode(2).Text
        End If
    End If
End Sub

Private Sub cmbConvert2_Click()
If glbCompSerial = "S/N - 2344W" Then 'Ticket #27356 Franks 10/27/2015
    txtConvert(1).Text = Left(cmbConvert2.Text, 1)
End If
End Sub

Private Sub cmbConvert3_Click()
If glbCompSerial = "S/N - 2442W" Then  'Pacific Sands Ticket #22352 Franks 01/11/2013
    txtConvert(2).Text = Left(cmbConvert3.Text, 1)
End If
'If glbCompSerial = "S/N - 2417W" Then  'County of Perth - Ticket #24497 Franks 10/23/2013
If NewAccpacMatrix Then
    txtConvert(2).Text = Left(cmbConvert3.Text, 1)
End If
If glbCompSerial = "S/N - 2436W" Then  'Family Day  - Ticket #23779 Franks 05/30/2013
    txtConvert(2).Text = cmbConvert3.Text
End If
End Sub

Private Sub cmbConvert4_Click()
If glbCompSerial = "S/N - 2436W" Then  'Family Day  - Ticket #21152 Franks 11/22/2012
    txtConvert(3).Text = getTxtCon3FamilyDay(cmbConvert4.Text)
ElseIf glbCompSerial = "S/N - 2437W" Then  'KN&V  - Ticket #21096 Franks 12/18/2012
    txtConvert(3).Text = getTxtCon3KNV(cmbConvert4.Text)
ElseIf glbCompSerial = "S/N - 2442W" Then  'Pacific Sands Ticket #22352 Franks 01/11/2013
    txtConvert(3).Text = getTxtCon3PacificSands(cmbConvert4.Text)
ElseIf glbCompSerial = "S/N - 2457W" Then  'McLeod Law - Ticket #24864 Franks 06/11/2014
    txtConvert(3).Text = getTxtCon3RegAccpac(cmbConvert4.Text)
ElseIf glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
    txtConvert(3).Text = getTxtCon3RegAccpac(cmbConvert4.Text)
'ElseIf glbCompSerial = "S/N - 2417W" Then  'County of Perth - Ticket #24497 Franks 10/23/2013
ElseIf NewAccpacMatrix Then
    txtConvert(3).Text = getTxtCon3RegAccpac(cmbConvert4.Text)
ElseIf NewPayWebMatrix Then 'Let's Talk Science Ticket #27072 10/14/2015
    txtConvert(3).Text = getTxtCon3RegPayWeb(cmbConvert4.Text)
Else
    txtConvert(3).Text = Left(cmbConvert4.Text, 1)
End If
End Sub
Private Function getTxtCon3KNV(xTxt) 'KN&V  - Ticket #21096 Franks 12/18/2012
Dim retval
    If Left(xTxt, 2) = "MC" Then
        retval = "MC"
    ElseIf Left(xTxt, 2) = "ME" Then
        retval = "ME"
    ElseIf Left(xTxt, 2) = "MB" Then
        retval = "MB"
    ElseIf Left(xTxt, 1) = "P" Then
        retval = "P"
    Else
        retval = ""
    End If
    getTxtCon3KNV = retval
End Function

Private Function getTxtCon3RegPayWeb(xTxt) 'Let's Talk Science Ticket #27072 10/14/2015
Dim retval
    If Left(xTxt, 4) = "C-24" Then
        retval = "C-24"
    ElseIf Left(xTxt, 4) = "E-24" Then
        retval = "E-24"
    ElseIf Left(xTxt, 4) = "B-24" Then
        retval = "B-24"
    ElseIf Left(xTxt, 4) = "C-26" Then
        retval = "C-26"
    ElseIf Left(xTxt, 4) = "E-26" Then
        retval = "E-26"
    ElseIf Left(xTxt, 4) = "B-26" Then
        retval = "B-26"
    ElseIf Left(xTxt, 4) = "M-ER" Then
        retval = "M-ER"
    ElseIf Left(xTxt, 4) = "M-EE" Then
        retval = "M-EE"
    ElseIf Left(xTxt, 1) = "P" Then
        retval = "P"
    ElseIf Left(xTxt, 1) = "N" Then
        retval = "N"
    Else
        retval = ""
    End If
    getTxtCon3RegPayWeb = retval
End Function

Private Function getTxtCon3RegAccpac(xTxt) 'County of Perth - Ticket #24497 Franks 10/23/2013
Dim retval
    If Left(xTxt, 4) = "C-24" Then
        retval = "C-24"
    ElseIf Left(xTxt, 4) = "E-24" Then
        retval = "E-24"
    ElseIf Left(xTxt, 4) = "B-24" Then
        retval = "B-24"
    ElseIf Left(xTxt, 4) = "C-26" Then
        retval = "C-26"
    ElseIf Left(xTxt, 4) = "E-26" Then
        retval = "E-26"
    ElseIf Left(xTxt, 4) = "B-26" Then
        retval = "B-26"
    ElseIf Left(xTxt, 4) = "M-ER" Then
        retval = "M-ER"
    ElseIf Left(xTxt, 4) = "M-EE" Then
        retval = "M-EE"
    ElseIf Left(xTxt, 1) = "P" Then
        retval = "P"
    ElseIf Left(xTxt, 1) = "N" Then
        retval = "N"
    Else
        retval = ""
    End If
    getTxtCon3RegAccpac = retval
End Function

Private Function getTxtCon3PacificSands(xTxt) 'Pacific Sands Ticket #22352 Franks 01/10/2012
Dim retval
    If Left(xTxt, 4) = "C-26" Then
        retval = "C-26"
    ElseIf Left(xTxt, 4) = "E-26" Then
        retval = "E-26"
    ElseIf Left(xTxt, 1) = "P" Then
        retval = "P"
    ElseIf Left(xTxt, 1) = "N" Then
        retval = "N"
    Else
        retval = ""
    End If
    getTxtCon3PacificSands = retval
End Function
Private Function getTxtCon3FamilyDay(xTxt) 'Family Day  - Ticket #21152 Franks 11/22/2012
Dim retval
    If Left(xTxt, 4) = "C-24" Then
        retval = "C-24"
    ElseIf Left(xTxt, 4) = "E-24" Then
        retval = "E-24"
    ElseIf Left(xTxt, 1) = "P" Then
        retval = "P"
    Else
        retval = ""
    End If
    getTxtCon3FamilyDay = retval
End Function

Private Sub cmbType_Change()
If cmbType.Text = "DEPT" Then
    clpCode(0).LookupType = 2
Else
    clpCode(0).LookupType = 0
    Select Case cmbType.Text
        Case "ATTD": clpCode(0).TablName = "ADRE"
        Case "BENE": clpCode(0).TablName = "BNCD"
        Case "EARN": clpCode(0).TablName = "EARN"
        Case "DOLE": clpCode(0).TablName = "EDOL"
        Case "PAYT": clpCode(0).TablName = "PYTC"
        Case "UNIO": clpCode(0).TablName = "EDOR"
    End Select
    
    'Ticket #26247 - if Benefit Code then allow 10 char
    If cmbType = "BENE" Then
        clpCode(0).MaxLength = 10
    End If
    'Ticket #28349 Franks 03/31/2016 - Attendance Code 4 chars
    If cmbType = "ATTD" Then
        clpCode(0).MaxLength = 4
    End If
    
End If
Call INI_Controls(Me)
'clpCode(0).RefreshDescription

Call Chglabels

End Sub
Private Sub Chglabels()
    If (glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2292W") Then 'Namasco 'Country of Elgin
       If clpCode(0).TablName = "PYTC" Then
            lblTitle(2).Caption = "Excel Cell" 'Convert 1
            lblTitle(3).Caption = "Earnings/Deduction Indicator(E or D)" 'Convert 2
            lblTitle(4).Caption = "Company Code" 'Convert 3
            lblTitle(6).Caption = "Rate Factor" 'Convert Number 1
            lblTitle(7).Caption = "Rate Maximum" 'Convert Number 2
            txtConvert(3).Enabled = False
            lblTitle(5).Enabled = False
            lblDept.Enabled = False
            clpDept.Enabled = False
        Else
            lblTitle(2).Caption = "Convert 1" 'Convert 1
            lblTitle(3).Caption = "Convert 2" 'Convert 2
            lblTitle(4).Caption = "Convert 3" 'Convert 1
            lblTitle(6).Caption = "Convert Number 1" 'Convert Number 1
            lblTitle(7).Caption = "Convert Number 2" 'Convert Number 2
            txtConvert(3).Enabled = True
            lblTitle(5).Enabled = True
            lblDept.Enabled = True
            clpDept.Enabled = True
        End If
    End If
    If clpCode(0).TablName = "PYTC" Then
        lblDept.Enabled = False
        clpDept.Enabled = False
        If glbCompSerial = "S/N - 2431W" Or glbWFC Then 'Ticket #21638 - BACI Franks 10/01/2012
            lblTitle(2).Caption = "Payroll Code"
        Else
            lblTitle(2).Caption = "Excel Cell" 'Convert 1
        End If
        'lblTitle(3).Caption = "Earnings/Deduction Indicator(E or D)" 'Convert 2
        If glbCompSerial = "S/N - 2431W" Then 'BACI Ticket #22736 Franks 11/01/2012
            lblTitle(3).Caption = "Earnings/Deduction/Hours(E, D or H)"
        Else
            lblTitle(3).Caption = "Earnings/Deduction/Memo(E, D or M)" 'Convert 2
        End If
        If glbWFC Then 'Ticket #22089 06/11/2012
            lblTitle(4).Caption = "NGS/ADP"
        Else
            lblTitle(4).Caption = "Company Code" 'Convert 3
        End If
        lblTitle(5).Enabled = False
        txtConvert(3).Enabled = False
        lblTitle(6).Caption = "Rate Factor" 'Convert Number 1
        lblTitle(7).Caption = "Rate Maximum" 'Convert Number 2
    Else
        lblTitle(2).Caption = "Convert 1" 'Convert 1
        lblTitle(3).Caption = "Convert 2" 'Convert 2
        lblTitle(4).Caption = "Convert 3" 'Convert 1
        lblTitle(6).Caption = "Convert Number 1" 'Convert Number 1
        lblTitle(7).Caption = "Convert Number 2" 'Convert Number 2
        txtConvert(3).Enabled = True
        lblTitle(5).Enabled = True
        lblDept.Enabled = True
        clpDept.Enabled = True
    End If
    
    If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #20696 Franks 10/14/2011
        If cmbType.Text = "BENE" Then
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblNote2.Caption = "PST Rate for M - Monthly Amount"
            lblNote2.Visible = True
        Else
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
        End If
    End If
        
    If glbCompSerial = "S/N - 2439W" Then  'OK Tire  - Ticket #21519 Franks 06/14/2012
        If cmbType.Text = "BENE" Then
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblNote2.Caption = "PST Rate "
            lblNote2.Visible = True
        Else
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
        End If
    End If
    If glbCompSerial = "S/N - 2436W" Then  'Family Day  - Ticket #21152 Franks 11/22/2012
        If cmbType.Text = "BENE" Then
            'Ticket #23779 Franks 05/30/2013
            'convert 3
            cmbConvert3.Left = txtConvert(2).Left
            cmbConvert3.Top = txtConvert(2).Top
            lblTitle(4).FontBold = True
            cmbConvert3.Visible = True
            
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            'lblNote2.Caption = "PST Rate "
            'lblNote2.Visible = True
        Else
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
        End If
    End If
    
    If glbCompSerial = "S/N - 2410W" Then  'Frontenac  - Ticket #25122 Franks 03/07/2014
        If cmbType.Text = "UNIO" Then
            lblTitle(2).Caption = "Benefit Group"
        Else
            lblTitle(2).Caption = "Convert 1"
        End If
    End If
    
    'If glbCompSerial = "S/N - 2417W" Then  'County of Perth - Ticket #24497 Franks 10/23/2013
    If NewAccpacMatrix Then
        Call ScreenSetupCountyPerth
    End If
    
    If NewPayWebMatrix Then 'Let's Talk Science Ticket #27072 10/14/2015
        Call ScreenSetupNewPayWeb
    End If
    
    If glbCompSerial = "S/N - 2437W" Then  'KN&V  - Ticket #21096 Franks 12/18/2012
        If cmbType.Text = "BENE" Then
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            'lblNote2.Caption = "PST Rate "
            'lblNote2.Visible = True
        Else
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
        End If
    End If
    
    If glbCompSerial = "S/N - 2430W" Then 'kidslink Ticket #21858 Franks 04/04/2012
        If cmbType.Text = "ATTD" Then
            lblTitle(2).Caption = "Accpac Earning Code"
        Else
            lblTitle(2).Caption = "Convert 1"
        End If
    End If

    If glbCompSerial = "S/N - 2411W" Then 'WDGPH Ticket #27978 Franks 05/16/2016
        If cmbType.Text = "ATTD" Then
            lblTitle(2).Caption = "Payroll Code"
            'lblTitle(3).Caption = "Exclude FT/PT"
            'Ticket #29148 Franks 02/28/2017 - do not use it
            'Ticket #30150 Franks 06/20/2017 - new logics
            'lblTitle(3).Caption = "CA/TR/TRP/TR<1/TP<1" ' "Convert 2"
            lblTitle(3).Caption = "CA/TR/TRP/TF12/TP12" ' "Convert 2"
            lblTitle(4).Caption = "FT/PT" ' "Convert 3"
            lblTitle(5).Caption = "All Employees" ' "Convert 4"
        Else
            lblTitle(2).Caption = "Convert 1"
            lblTitle(3).Caption = "Convert 2"
            lblTitle(4).Caption = "Convert 3"
            lblTitle(5).Caption = "Convert 4"
        End If
    End If

    If glbCompSerial = "S/N - 2442W" Then  'Pacific Sands  - Ticket #22352 Franks 01/11/2013
        If cmbType.Text = "BENE" Then
            lblTitle(2).FontBold = True
            'convert 3
            cmbConvert3.Left = txtConvert(2).Left
            cmbConvert3.Top = txtConvert(2).Top
            lblTitle(4).FontBold = True
            cmbConvert3.Visible = True
            'convert 4
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblTitle(2).Caption = "Pay Code"
            lblTitle(3).Caption = "Distcode"
            lblTitle(4).Caption = "Category"
            lblTitle(5).Caption = "Amount"
        Else
            lblTitle(2).FontBold = False
            lblTitle(4).FontBold = False
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
            lblTitle(2).Caption = "Convert 1"
            lblTitle(3).Caption = "Convert 2"
            lblTitle(4).Caption = "Convert 3"
            lblTitle(5).Caption = "Convert 4"
        End If
    End If

    If glbCompSerial = "S/N - 2457W" Then  'McLeod Law - Ticket #24864 Franks 06/11/2014
        If cmbType.Text = "BENE" Then
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(2).FontBold = True
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblTitle(2).Caption = "Pay Code"
            lblTitle(5).Caption = "Amount"
            'Ticket #24864 Franks 11/11/2014
            clpCode(2).Visible = True
            clpCode(2).Top = txtConvert(1).Top
            clpCode(2).Left = clpCode(0).Left
            lblTitle(3).Caption = lStr("Administered By")
            lblTitle(3).FontBold = True
        Else
            lblTitle(2).FontBold = False
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
            lblTitle(2).Caption = "Convert 1"
            lblTitle(5).Caption = "Convert 4"
            lblTitle(3).Caption = "Convert 2" 'Ticket #24864 Franks 11/11/2014
            lblTitle(3).FontBold = False
        End If
    End If
    
    If glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
        If cmbType.Text = "BENE" Then
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(2).FontBold = True
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblTitle(2).Caption = "Pay Code"
            lblTitle(5).Caption = "Amount"
            'Ticket #24864 Franks 11/11/2014
            clpCode(2).Visible = True
            clpCode(2).Top = txtConvert(1).Top
            clpCode(2).Left = clpCode(0).Left
            lblTitle(3).Caption = lStr("Administered By")
            lblTitle(3).FontBold = True
        Else
            lblTitle(2).FontBold = False
            lblTitle(5).FontBold = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
            lblTitle(2).Caption = "Convert 1"
            lblTitle(5).Caption = "Convert 4"
            lblTitle(3).Caption = "Convert 2" 'Ticket #24864 Franks 11/11/2014
            lblTitle(3).FontBold = False
        End If
    End If
    
    'Ticket #26247 - if Benefit Code then allow 10 char
    If cmbType = "BENE" Then
        clpCode(0).MaxLength = 10
    End If
    'Ticket #28349 Franks 03/31/2016 - Attendance Code 4 chars
    If cmbType = "ATTD" Then
        clpCode(0).MaxLength = 4
    End If
    
    If glbCompSerial = "S/N - 2344W" Then 'Ticket #27356 Franks 10/27/2015
        If cmbType.Text = "ATTD" Then
            lblTitle(2).Caption = "Pay Code"
            lblTitle(3).Caption = "Affecting Standard Hours"
            cmbConvert2.Left = txtConvert(1).Left
            cmbConvert2.Top = txtConvert(1).Top
            cmbConvert2.Visible = True
        Else
            lblTitle(2).FontBold = False
            lblTitle(5).FontBold = False
            cmbConvert2.Visible = False
            lblNote2.Visible = False
            lblTitle(2).Caption = "Convert 1"
            lblTitle(5).Caption = "Convert 4"
            lblTitle(3).Caption = "Convert 2" 'Ticket #24864 Franks 11/11/2014
            lblTitle(3).FontBold = False
        End If
    End If
    
End Sub

Private Sub cmbType_Click()
If UpdateState = OPENING Or UpdateState = NewRecord Then
'If cmdOK.Enabled Then
    If cmbType = "DEPT" Then
        clpCode(0).Enabled = False
        clpCode(0) = ""
        lblTitle(1).Font.Bold = False
        clpDept.Enabled = True
        lblDept.Font.Bold = True
        txtConvert(0).MaxLength = 7
        txtConvert(1).MaxLength = 7
        txtConvert(2).MaxLength = 7
        txtConvert(3).MaxLength = 7
    Else
        clpCode(0).Enabled = True
        lblTitle(1).Font.Bold = True
        clpDept.Enabled = True
        lblDept.Font.Bold = False
        
        'Hemu - changing to 10 chrs because in other part of the code it has been set to 10, so I am
        'assuming whoever made that change forgot to make the change in this part of the code. Since
        'there was no serial # control I am assuming it is for all.
        txtConvert(0).MaxLength = 10    '4
        txtConvert(1).MaxLength = 10    '4
        txtConvert(2).MaxLength = 10    '4
        txtConvert(3).MaxLength = 10    '4
    End If
    
End If
Call cmbType_Change

txtTransfer.Text = cmbType.Text
'clpCode(0).TABLName = cmbType.Text


End Sub

Private Sub cmbType_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmbType_LostFocus()
    'txtTransfer.Text = cmbType.Text
    '  clpCode(0).TABLName = cmbType.Text
End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If


'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
rsDATA.CancelUpdate
Call Display_Value


'Call ST_UPD_MODE(True) ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRMATRIX", "Cancel")
Call RollBack '09June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim x As Integer
Dim xID As Integer

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
If vbxTrueGrid.SelBookmarks.count > 1 Then
    Msg = Msg & "These Records?"
Else
    Msg = Msg & "This Record?"
End If
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'7.9 Enhancement to delete multiple records.
If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add Data1.Recordset.Bookmark
For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
    Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(x)
    xID = Data1.Recordset("M_ID")
    
    gdbAdoIhr001.BeginTrans
    'rsDATA.Delete
    gdbAdoIhr001.Execute "DELETE FROM HRMATRIX WHERE M_ID=" & xID
    gdbAdoIhr001.CommitTrans
    DoEvents
Next

Data1.Refresh


Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRMATRIX", "Delete")
Call RollBack '09June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
cmbType.SetFocus
Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRMATRIX", "Modify")
Call RollBack '09June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

lblCNum.Caption = "001"
'txtNumConvert(1).Text = ""
'txtNumConvert(2).Text = ""

fglbNew = True
Call SET_UP_MODE

'Call ST_UPD_MODE(True)
cmbType.SetFocus
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRMATRIX", "Add")
Call RollBack '09June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x%
Dim bmk As Variant

On Error GoTo cmdOK_Err
If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

txtTransfer.Text = cmbType.Text
Select Case cmbType.Text
    Case "ATTD": txtType = "ADRE"
    Case "BENE": txtType = "BNCD"
    Case "EARN": txtType = "EARN"
    Case "DOLE": txtType = "EDOL"
    Case "DEPT": txtType = "DEPT"
    Case "PAYT": txtType = "PAYT"
    Case "UNIO": txtType = "UNIO"
End Select
If Not chkEPayroll() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
If glbCompSerial = "S/N - 2390W" Then 'Ticket #16311
    If cmbType.Text = "DEPT" And UCase(txtConvert(0).Text) = "Y" Then
        Call Pause(1)
        Call Codes_Master_Integration("DEPT", clpDept.Text)
    End If
End If
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRMATRIX", "Update")
Call RollBack '09June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Payroll Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub
Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Payroll Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub combType()
cmbType.AddItem "ATTD"
cmbType.AddItem "BENE"
cmbType.AddItem "EARN"
cmbType.AddItem "DOLE"
cmbType.AddItem "DEPT"
cmbType.AddItem "PAYT"
cmbType.AddItem "UNIO"

If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #20696 Franks 10/14/2011
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "M - Monthly Amount"
    cmbConvert4.Visible = True
End If

If glbCompSerial = "S/N - 2439W" Then  'OK Tire  - Ticket #21519 Franks 06/14/2012
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "M - Monthly Amount"
    cmbConvert4.AddItem "A - Annual Amount"
    'cmbConvert4.Visible = True
End If
If glbCompSerial = "S/N - 2436W" Then  'Family Day  - Ticket #21152 Franks 11/22/2012
    'Ticket #23779 Franks 05/30/2013
    cmbConvert3.Clear
    cmbConvert3.AddItem "0019"
    cmbConvert3.AddItem "1611"
    cmbConvert3.Width = 1215
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "C-24 - Use Annual Company Cost divided by 24"
    cmbConvert4.AddItem "E-24 - Use Annual Employee Cost divided by 24"
    cmbConvert4.Width = 3800
End If
If glbCompSerial = "S/N - 2437W" Then  'KN&V  - Ticket #21096 Franks 12/18/2012
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "MC - Monthly Company"
    cmbConvert4.AddItem "ME - Monthly Employee"
    cmbConvert4.AddItem "MB - Monthly Company and Monthly Employee"
    cmbConvert4.Width = 3800
End If

If glbCompSerial = "S/N - 2442W" Then  'Pacific Sands  - Ticket #22352 Franks 01/11/2013
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "C-26 - Use Annual Company Cost divided by 26"
    cmbConvert4.AddItem "E-26 - Use Annual Employee Cost divided by 26"
    cmbConvert4.AddItem "N - No $ are transferred"
    cmbConvert4.Width = 3800
    
    cmbConvert3.Clear
    cmbConvert3.AddItem "1 - Accrual"
    cmbConvert3.AddItem "2 - Earning"
    cmbConvert3.AddItem "3 - Advance"
    cmbConvert3.AddItem "4 - Deduction"
    cmbConvert3.AddItem "5 - Expense Reimbursement"
    cmbConvert3.AddItem "6 - Benefit"
    cmbConvert3.Width = 3800
End If

If glbCompSerial = "S/N - 2457W" Then  'McLeod Law - Ticket #24864 Franks 06/11/2014
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "C-24 - Use Annual Company Cost divided by 24"
    cmbConvert4.AddItem "E-24 - Use Annual Employee Cost divided by 24"
    cmbConvert4.Width = 3800
End If

If glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "C-24 - Use Annual Company Cost divided by 24"
    cmbConvert4.AddItem "E-24 - Use Annual Employee Cost divided by 24"
    cmbConvert4.Width = 3800
End If

'If glbCompSerial = "S/N - 2417W" Then  'County of Perth - Ticket #24497 Franks 10/23/2013
If NewAccpacMatrix Then
    cmbList4Accpac
End If

If NewPayWebMatrix Then  'Let's Talk Science Ticket #27072 10/14/2015
    cmbList4PayWeb
End If

If glbCompSerial = "S/N - 2344W" Then 'Ticket #27356 Franks 10/27/2015
    cmbConvert2.Clear
    cmbConvert2.AddItem "1 - Export Hours without Affecting Standard Hours"
    cmbConvert2.AddItem "2 - Deducts from Hours per Pay Period"
    cmbConvert2.AddItem "3 - Extra Hours Exceeding per Pay Period"
End If

End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "PAYROLL", "SELECT")

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

'Ticket #24267 Franks 01/14/2014
'2257 - HCCAS Ticket #24452 Franks 01/27/2014
'2449W PICS Ticket #23288 Franks 02/25/2014
'2466W Chiefs of Ontario Ticket #25879 Franks 09/22/2014
If glbCompSerial = "S/N - 2417W" Or glbCompSerial = "S/N - 2331W" Or glbCompSerial = "S/N - 2257W" Or glbCompSerial = "S/N - 2391W" Or glbCompSerial = "S/N - 2449W" Or glbCompSerial = "S/N - 2466W" Then
    NewAccpacMatrix = True
ElseIf glbCompSerial = "S/N - 2376W" Then 'Ticket #29946 Franks 03/15/2017
    NewAccpacMatrix = True
ElseIf glbCompSerial = "S/N - 2489W" Then 'Ticket #29568 Franks 06/07/2017 - Sudbury & District Health Unit
    NewAccpacMatrix = True
ElseIf glbCompSerial = "S/N - 2485W" Then 'Ticket #28652 Franks 06/28/2017 - Mississaugas of Scugog Island First Nation
    NewAccpacMatrix = True
Else
    NewAccpacMatrix = False
End If

'Let's Talk Science Ticket #27072 10/14/2015
If glbCompSerial = "S/N - 2353W" Then
    NewPayWebMatrix = True
Else
    NewPayWebMatrix = False
End If

If glbCompSerial = "S/N - 2489W" Then 'Ticket #29568 Franks 08/21/2017 - Sudbury & District Health Unit
    lblSection.Caption = lStr("Union")
    vbxTrueGrid.Columns(0).Caption = lStr("Union")
    clpCode(1).TablName = "EDOR"
Else
    vbxTrueGrid.Columns(0).Caption = lStr("Section")
End If

Me.Show
glbOnTop = "FRMPAYROLL"
Call combType
Screen.MousePointer = HOURGLASS

If glbCompSerial = "S/N - 2430W" Then 'Ticket #21167 Franks 11/07/2011
    lblDept.Caption = "Program"
    clpDept.LookupType = ProjectCode
    vbxTrueGrid.Columns(3).Caption = "Program"
End If


'Ticket #13035
If ConNume Then
    txtNumConvert(1).DataField = "M_USER_NUM_1"
    txtNumConvert(2).DataField = "M_USER_NUM_2"
End If

Data1.ConnectionString = glbAdoIHRDB
If glbWFC Then
    SQLQ = "SELECT * FROM HRMATRIX "
    If Len(glbPlantCode) > 0 Then
        SQLQ = SQLQ & "WHERE M_SECTION = '" & glbPlantCode & "' "
    End If
    SQLQ = SQLQ & "ORDER BY M_SECTION,M_DEFTYPE,M_CODE  "
    Data1.RecordSource = SQLQ
Else
    If glbCompSerial = "S/N - 2382W" Then 'Namasco
        Data1.RecordSource = "SELECT * FROM HRMATRIX ORDER BY M_DEFTYPE,M_CONVERT3,M_CODE "
    Else
        Data1.RecordSource = "SELECT * FROM HRMATRIX ORDER BY M_DEFTYPE,M_CODE "
    End If
End If
Data1.Refresh
'Frank May 17,2002
Call setRptCaption(Me)
Screen.MousePointer = DEFAULT
'Call Display_Value
Call ST_UPD_MODE(False)
If Not gSec_Matrix Then                                     'May99 js
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If                                                  '
'vbxTrueGrid.Columns(0).Caption = lStr("Section")
'Ticket #16392 US Payroll not need Plant Code
'If glbWFC Then
'    lblSection.FontBold = True
'End If
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT                           '
End Sub

Private Function ConNume()
Dim rsTTemp As New ADODB.Recordset
Dim SQLQ As String
On Error GoTo Err_Line
    ConNume = False
    SQLQ = "SELECT M_USER_NUM_1 FROM HRMATRIX"
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    rsTTemp.Close
    ConNume = True
    Exit Function
Err_Line:

End Function

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
cmbType.Enabled = TF            '
clpCode(0).Enabled = TF
clpDept.Enabled = TF                    '
txtConvert(0).Enabled = TF      '
txtConvert(1).Enabled = TF      '
txtConvert(2).Enabled = TF      '
txtConvert(3).Enabled = TF      '
If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
 '   cmdDelete.Enabled = False
End If
'Frank May 17,2002
If TF Then
    If cmbType = "DEPT" Then
        clpCode(0).Enabled = False
        lblTitle(1).Font.Bold = False
        clpDept.Enabled = True
        lblDept.Font.Bold = True
        txtConvert(0).MaxLength = 7
        txtConvert(1).MaxLength = 7
        txtConvert(2).MaxLength = 7
        txtConvert(3).MaxLength = 7
    Else
        clpCode(0).Enabled = True
        
        'If cmbType = "ATTD" Then
        '    clpCode(0).MaxLength = 4
        'End If
        'Ticket #20484 Franks 06/17/2011
        clpCode(0).MaxLength = 7

        'Ticket #26247 - if Benefit Code then allow 10 char
        If cmbType = "BENE" Then
            clpCode(0).MaxLength = 10
        End If
        'Ticket #28349 Franks 03/31/2016 - Attendance Code 4 chars
        If cmbType = "ATTD" Then
            clpCode(0).MaxLength = 4
        End If
        
        lblTitle(1).Font.Bold = True
        clpDept.Enabled = True
        lblDept.Font.Bold = False 'True
        txtConvert(0).MaxLength = 10 '4
        txtConvert(1).MaxLength = 10 '4
        txtConvert(2).MaxLength = 10 '4
        txtConvert(3).MaxLength = 10 '4
    End If
End If
'Frank May 17,2002
End Sub

Private Sub txtConvert_Change(Index As Integer)
If glbCompSerial = "S/N - 2382W" Then  'Samuel - Ticket #20696 Franks 10/14/2011
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "M" Then
            cmbConvert4.ListIndex = 1
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If

If glbCompSerial = "S/N - 2439W" Then  'OK Tire  - Ticket #21519 Franks 06/14/2012
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "M" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "A" Then
            cmbConvert4.ListIndex = 2
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If
If glbCompSerial = "S/N - 2436W" Then  'Family Day  - Ticket #21152 Franks 11/22/2012
    If Index = 2 Then 'Ticket #23779 Franks 05/30/2013
        If txtConvert(2).Text = "0019" Then
            cmbConvert3.ListIndex = 0
        ElseIf txtConvert(2).Text = "1611" Then
            cmbConvert3.ListIndex = 1
        End If
    End If
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "C-24" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "E-24" Then
            cmbConvert4.ListIndex = 2
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If
If glbCompSerial = "S/N - 2437W" Then  'KN&V  - Ticket #21096 Franks 12/18/2012
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "MC" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "ME" Then
            cmbConvert4.ListIndex = 2
        ElseIf txtConvert(3).Text = "MB" Then
            cmbConvert4.ListIndex = 3
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If
If glbCompSerial = "S/N - 2442W" Then  'Pacific Sands Ticket #22352 Franks 01/11/2013
    If Index = 2 Then
        If txtConvert(2).Text = "1" Then
            cmbConvert3.ListIndex = 0
        ElseIf txtConvert(2).Text = "2" Then
            cmbConvert3.ListIndex = 1
        ElseIf txtConvert(2).Text = "3" Then
            cmbConvert3.ListIndex = 2
        ElseIf txtConvert(2).Text = "4" Then
            cmbConvert3.ListIndex = 3
        ElseIf txtConvert(2).Text = "5" Then
            cmbConvert3.ListIndex = 4
        ElseIf txtConvert(2).Text = "6" Then
            cmbConvert3.ListIndex = 5
        Else
            cmbConvert3.ListIndex = -1
        End If
    End If
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "C-26" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "E-26" Then
            cmbConvert4.ListIndex = 2
        ElseIf txtConvert(3).Text = "N" Then
            cmbConvert4.ListIndex = 3
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If

If glbCompSerial = "S/N - 2457W" Then  'McLeod Law - Ticket #24864 Franks 06/11/2014
    If Index = 1 Then 'Ticket #24864 Franks 11/11/2014
        clpCode(2).Text = txtConvert(1).Text
    End If
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "C-24" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "E-24" Then
            cmbConvert4.ListIndex = 2
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If

If glbCompSerial = "S/N - 2475W" Then   'Ticket #27436 - Super Channel
    If Index = 1 Then 'Ticket #24864 Franks 11/11/2014
        clpCode(2).Text = txtConvert(1).Text
    End If
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "C-24" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "E-24" Then
            cmbConvert4.ListIndex = 2
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End If

'If glbCompSerial = "S/N - 2417W" Then  'County of Perth - Ticket #24497 Franks 10/23/2013
If NewAccpacMatrix Then
    Call DispList4Accpac(Index)
End If

If NewPayWebMatrix Then 'Let's Talk Science Ticket #27072 10/14/2015
    Call DispList4PayWeb(Index)
End If

If glbCompSerial = "S/N - 2344W" Then 'Ticket #27356 Franks 10/27/2015
    If Index = 1 Then
        If txtConvert(1).Text = "1" Then
            cmbConvert2.ListIndex = 0
        ElseIf txtConvert(1).Text = "2" Then
            cmbConvert2.ListIndex = 1
        ElseIf txtConvert(1).Text = "3" Then
            cmbConvert2.ListIndex = 2
        Else
            cmbConvert2.ListIndex = -1
        End If
    End If
End If

End Sub

Private Sub txtConvert_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub


'Private Sub  clpDept_Change()
'End Sub
'Private Sub txtDept_DblClick()
'    Call Get_Dept(False)
'    txtDept.Text = glbDept
'    lblDeptDesc.Caption = glbDeptDesc
'End Sub
'Private Sub txtDept_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDept_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

Private Sub txtTransfer_Change()

If txtTransfer = "ATTD" Then
    cmbType.ListIndex = 0
End If

If txtTransfer = "BENE" Then
    cmbType.ListIndex = 1
End If

If txtTransfer = "EARN" Then
    cmbType.ListIndex = 2
End If

If txtTransfer = "DOLE" Then
    cmbType.ListIndex = 3
End If
'Frank May 17,2002
If txtTransfer = "DEPT" Then
    cmbType.ListIndex = 4
End If
'Frank May 17,2002
If txtTransfer = "PAYT" Then
    cmbType.ListIndex = 5
End If
If txtTransfer = "UNIO" Then
    cmbType.ListIndex = 6
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
        
        If glbWFC Then
            SQLQ = "SELECT * FROM HRMATRIX "
            If Len(glbPlantCode) > 0 Then
                SQLQ = SQLQ & "WHERE M_SECTION = '" & glbPlantCode & "' "
            End If
        Else
            SQLQ = "SELECT * FROM HRMATRIX "
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

If txtTransfer = "ATTD" Then
    I% = 0
    txtType = "ADRE"
End If

If txtTransfer = "BENE" Then
    I% = 1
    txtType = "BNCD"
End If

If txtTransfer = "EARN" Then
    I% = 2
    txtType = "EARN"
End If

If txtTransfer = "DOLE" Then
    I% = 3
    txtType = "EDOL"
End If

cmbType.ListIndex = I%


Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "Payroll Matrix", "Select")
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

   
''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRMATRIX where M_ID= " & Data1.Recordset!M_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
    If rsDATA("M_TYPE") = "BNCD" Then clpCode(0).MaxLength = 10
    If rsDATA("M_TYPE") = "ADRE" Then clpCode(0).MaxLength = 4
    
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    
    Call Chglabels
    
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
UpdateRight = gSec_Matrix
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

Private Sub ScreenSetupNewPayWeb() 'Let's Talk Science Ticket #27072 10/14/2015
        If cmbType.Text = "BENE" Then
            lblTitle(2).FontBold = True
            'convert 3
            'cmbConvert3.Left = txtConvert(2).Left
            'cmbConvert3.Top = txtConvert(2).Top
            'lblTitle(4).FontBold = True
            'cmbConvert3.Visible = True
            'convert 4
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblTitle(2).Caption = "Pay Code"
            'lblTitle(3).Caption = "Distcode"
            'lblTitle(4).Caption = "Category"
            lblTitle(5).Caption = "Amount"
            If glbCompSerial = "S/N - 2353W" Then
                lblTitle(6).Caption = "Add Ontario Tax %"
            Else
                lblTitle(6).Caption = "Add Tax %"
            End If
            lblTitle(6).FontBold = True
        Else
            lblTitle(2).FontBold = False
            lblTitle(4).FontBold = False
            lblTitle(5).FontBold = False
            cmbConvert3.Visible = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
            lblTitle(2).Caption = "Convert 1"
            lblTitle(3).Caption = "Convert 2"
            lblTitle(4).Caption = "Convert 3"
            lblTitle(5).Caption = "Convert 4"
            lblTitle(6).Caption = "Convert Number 1"
            lblTitle(6).FontBold = False
        End If

End Sub

Private Sub ScreenSetupCountyPerth() 'County of Perth - Ticket #24497 Franks 10/23/2013
        If cmbType.Text = "BENE" Then
            lblTitle(2).FontBold = True
            'convert 3
            cmbConvert3.Left = txtConvert(2).Left
            cmbConvert3.Top = txtConvert(2).Top
            lblTitle(4).FontBold = True
            cmbConvert3.Visible = True
            'convert 4
            cmbConvert4.Left = txtConvert(3).Left
            cmbConvert4.Top = txtConvert(3).Top
            lblTitle(5).FontBold = True
            cmbConvert4.Visible = True
            lblTitle(2).Caption = "Pay Code"
            lblTitle(3).Caption = "Distcode"
            lblTitle(4).Caption = "Category"
            lblTitle(5).Caption = "Amount"
        Else
            lblTitle(2).FontBold = False
            lblTitle(4).FontBold = False
            lblTitle(5).FontBold = False
            cmbConvert3.Visible = False
            cmbConvert4.Visible = False
            lblNote2.Visible = False
            lblTitle(2).Caption = "Convert 1"
            lblTitle(3).Caption = "Convert 2"
            lblTitle(4).Caption = "Convert 3"
            lblTitle(5).Caption = "Convert 4"
        End If
End Sub

Private Sub cmbList4PayWeb() 'Let's Talk Science Ticket #27072 10/14/2015
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "M-ER - Use Monthly Company Cost"
    cmbConvert4.AddItem "M-EE - Use Monthly Employee Cost"
    cmbConvert4.AddItem "C-24 - Use Annual Company Cost divided by 24"
    cmbConvert4.AddItem "E-24 - Use Annual Employee Cost divided by 24"
    cmbConvert4.AddItem "B-24 - Use both Annual Company and Annual Employee divided by 24"
    cmbConvert4.AddItem "C-26 - Use Annual Company Cost divided by 26"
    cmbConvert4.AddItem "E-26 - Use Annual Employee Cost divided by 26"
    cmbConvert4.AddItem "B-26 - Use both Annual Company and Annual Employee divided by 26"
    cmbConvert4.AddItem "N - No $ are transferred"
    cmbConvert4.Width = 5600
End Sub

Private Sub DispList4PayWeb(Index As Integer) 'Let's Talk Science Ticket #27072 10/14/2015
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "M-ER" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "M-EE" Then
            cmbConvert4.ListIndex = 2
        ElseIf txtConvert(3).Text = "C-24" Then
            cmbConvert4.ListIndex = 3
        ElseIf txtConvert(3).Text = "E-24" Then
            cmbConvert4.ListIndex = 4
        ElseIf txtConvert(3).Text = "B-24" Then
            cmbConvert4.ListIndex = 5
        ElseIf txtConvert(3).Text = "C-26" Then
            cmbConvert4.ListIndex = 6
        ElseIf txtConvert(3).Text = "E-26" Then
            cmbConvert4.ListIndex = 7
        ElseIf txtConvert(3).Text = "B-26" Then
            cmbConvert4.ListIndex = 8
        ElseIf txtConvert(3).Text = "N" Then
            cmbConvert4.ListIndex = 9
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End Sub

Private Sub cmbList4Accpac() 'County of Perth - Ticket #24497 Franks 10/23/2013
    cmbConvert3.Clear
    cmbConvert3.AddItem "1 - Accrual"
    cmbConvert3.AddItem "2 - Earning"
    cmbConvert3.AddItem "3 - Advance"
    cmbConvert3.AddItem "4 - Deduction"
    cmbConvert3.AddItem "5 - Expense Reimbursement"
    cmbConvert3.AddItem "6 - Benefit"
    cmbConvert3.Width = 2200
    
    cmbConvert4.Clear
    cmbConvert4.AddItem "P - Pay Period Amount"
    cmbConvert4.AddItem "M-ER - Use Monthly Company Cost"
    cmbConvert4.AddItem "M-EE - Use Monthly Employee Cost"
    cmbConvert4.AddItem "C-24 - Use Annual Company Cost divided by 24"
    cmbConvert4.AddItem "E-24 - Use Annual Employee Cost divided by 24"
    cmbConvert4.AddItem "B-24 - Use both Annual Company and Annual Employee divided by 24"
    cmbConvert4.AddItem "C-26 - Use Annual Company Cost divided by 26"
    cmbConvert4.AddItem "E-26 - Use Annual Employee Cost divided by 26"
    cmbConvert4.AddItem "B-26 - Use both Annual Company and Annual Employee divided by 26"
    cmbConvert4.AddItem "N - No $ are transferred"
    cmbConvert4.Width = 5600
End Sub

Private Sub DispList4Accpac(Index As Integer) 'County of Perth - Ticket #24497 Franks 10/23/2013
    If Index = 2 Then
        If txtConvert(2).Text = "1" Then
            cmbConvert3.ListIndex = 0
        ElseIf txtConvert(2).Text = "2" Then
            cmbConvert3.ListIndex = 1
        ElseIf txtConvert(2).Text = "3" Then
            cmbConvert3.ListIndex = 2
        ElseIf txtConvert(2).Text = "4" Then
            cmbConvert3.ListIndex = 3
        ElseIf txtConvert(2).Text = "5" Then
            cmbConvert3.ListIndex = 4
        ElseIf txtConvert(2).Text = "6" Then
            cmbConvert3.ListIndex = 5
        Else
            cmbConvert3.ListIndex = -1
        End If
    End If
    If Index = 3 Then
        If txtConvert(3).Text = "P" Then
            cmbConvert4.ListIndex = 0
        ElseIf txtConvert(3).Text = "M-ER" Then
            cmbConvert4.ListIndex = 1
        ElseIf txtConvert(3).Text = "M-EE" Then
            cmbConvert4.ListIndex = 2
        ElseIf txtConvert(3).Text = "C-24" Then
            cmbConvert4.ListIndex = 3
        ElseIf txtConvert(3).Text = "E-24" Then
            cmbConvert4.ListIndex = 4
        ElseIf txtConvert(3).Text = "B-24" Then
            cmbConvert4.ListIndex = 5
        ElseIf txtConvert(3).Text = "C-26" Then
            cmbConvert4.ListIndex = 6
        ElseIf txtConvert(3).Text = "E-26" Then
            cmbConvert4.ListIndex = 7
        ElseIf txtConvert(3).Text = "B-26" Then
            cmbConvert4.ListIndex = 8
        ElseIf txtConvert(3).Text = "N" Then
            cmbConvert4.ListIndex = 9
        Else
            cmbConvert4.ListIndex = -1
        End If
    End If
End Sub

