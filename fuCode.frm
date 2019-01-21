VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUCode 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Mass Table Master"
   ClientHeight    =   8490
   ClientLeft      =   525
   ClientTop       =   1470
   ClientWidth     =   11205
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNewCourseName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Tag             =   "00-Course Name"
      Top             =   7320
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   468
      Width           =   3135
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Course Code Replacement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   6
      Top             =   180
      Width           =   3135
   End
   Begin VB.TextBox txtCourseName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   16
      Tag             =   "00-Course Name"
      Top             =   6480
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Salary Distribution Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   3
      Top             =   1044
      Width           =   3735
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Position Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1332
      Width           =   3375
   End
   Begin INFOHR_Controls.CodeLookup clpNewCode 
      Height          =   285
      Left            =   2610
      TabIndex        =   14
      Top             =   5490
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      MaxLength       =   20
   End
   Begin INFOHR_Controls.CodeLookup clpOldCode 
      Height          =   285
      Left            =   2580
      TabIndex        =   11
      Top             =   4740
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      MaxLength       =   20
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      Height          =   285
      Left            =   1230
      TabIndex        =   10
      Top             =   4740
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin VB.CheckBox chkSep 
      Caption         =   "Split for Facilities"
      Height          =   345
      Left            =   1560
      TabIndex        =   12
      Top             =   5130
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Table Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1620
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "G/L Number Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   756
      Width           =   3735
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Division Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   468
      Width           =   3255
   End
   Begin VB.OptionButton optCodes 
      Caption         =   "Department Master"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3135
   End
   Begin VB.ComboBox cmbNewDIV 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   5520
      Width           =   795
   End
   Begin VB.TextBox txtTblName 
      Appearance      =   0  'Flat
      DataField       =   "TD_NAME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   4110
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid tblTables 
      Bindings        =   "fuCode.frx":0000
      Height          =   1995
      Left            =   120
      OleObjectBlob   =   "fuCode.frx":0014
      TabIndex        =   8
      Tag             =   "Tables Names Lookup"
      Top             =   2040
      Width           =   8955
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6240
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   8280
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   4080
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Data3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   15
      Tag             =   "00-Enter Section Code"
      Top             =   6000
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin MSAdodcLib.Adodc Data5 
      Height          =   330
      Left            =   120
      Top             =   7920
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
      Caption         =   "Data5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   2040
      Top             =   7920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Data4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1730
      TabIndex        =   17
      Tag             =   "00-Course Code"
      Top             =   6930
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ESCD"
      MaxLength       =   8
   End
   Begin VB.Label lblMsg 
      Caption         =   $"fuCode.frx":2B6E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   4200
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Label lblNewCourseName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Course Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   29
      Top             =   7357
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblCourseName 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   6525
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblCourseCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Course Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   120
      TabIndex        =   27
      Top             =   6967
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblVadim 
      Caption         =   "The Changes will not be transferred to Vadim"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4470
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   24
      Top             =   4470
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Facility"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1560
      TabIndex        =   23
      Top             =   4470
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblOldCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   795
   End
   Begin VB.Label lblNewCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   5145
      Width           =   750
   End
   Begin VB.Label lblTblNameDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unassigned"
      DataField       =   "TD_DESC"
      DataSource      =   "DATA1"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2700
      TabIndex        =   20
      Top             =   4140
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Table Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   4140
      Width           =   1035
   End
End
Attribute VB_Name = "frmUCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbState

Dim fglbSDate As Variant
Dim otxtDIV
Dim cleTable As New Collection, xItem() As New Collection
Dim RowMax As Integer
Dim ApptFacility()

Private Function chkFUComment()

Dim SQLQ As String, Msg$, dd&, Response%, x%
Dim DgDef As Variant, Title$, DCurPDate As Variant
Dim sCourseName As String
chkFUComment = False

'''On Error GoTo chkFUComment_Err

If optCodes(6).Value = True Then
    If txtCourseName.Text = "" Or Len(clpCode(1)) <= 0 Then
        MsgBox "You must enter Course Name and Course Code", vbCritical, "Error Updating Course Code"
        txtCourseName.SetFocus
        Exit Function
    End If
    If Len(clpCode(1)) > 0 And clpCode(1).Caption = "Unassigned" Then
       MsgBox "Invalid Course Code", vbCritical, "Course Code Error"
       clpCode(1).SetFocus
       Exit Function
    End If
    
    
    sCourseName = txtCourseName.Text
    SQLQ = "SELECT * From Term_HREDSEM"
    'Ticket #23369
    'SQLQ = SQLQ & " WHERE ES_COURSE = '" & Replace(txtCourseName, "'", "'+chr(39)+'") & "'"
    SQLQ = SQLQ & " WHERE UPPER(ES_COURSE) = '" & UCase(Replace(txtCourseName, "'", "''")) & "'"
    Data5.RecordSource = SQLQ
    Data5.Refresh
    SQLQ = "SELECT * From HREDSEM"
    'Ticket #23369
    'SQLQ = SQLQ & " WHERE ES_COURSE = '" & Replace(txtCourseName, "'", "'+chr(39)+'") & "'"
    SQLQ = SQLQ & " WHERE UPPER(ES_COURSE) = '" & UCase(Replace(txtCourseName, "'", "''")) & "'"
    Data4.RecordSource = SQLQ
    Data4.Refresh
    If Data4.Recordset.EOF Then
        MsgBox "Course Name does not exist", vbCritical, "Course Name Error"
        txtCourseName.SetFocus
        Exit Function
    End If

    chkFUComment = True
    Exit Function
End If

'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
If optCodes(7).Value = True Then
    If txtCourseName.Text = "" Then
        MsgBox "You must enter Old Course Name and New Course Name", vbCritical, "Error Updating Course Name"
        txtCourseName.SetFocus
        Exit Function
    End If
    If txtNewCourseName.Text = "" Then
        MsgBox "You must enter Old Course Name and New Course Name", vbCritical, "Error Updating Course Name"
        txtNewCourseName.SetFocus
        Exit Function
    End If
    
    'Create a recordset for Active Employees
    sCourseName = txtCourseName.Text
    SQLQ = "SELECT * From HREDSEM"
    SQLQ = SQLQ & " WHERE UPPER(ES_COURSE) = '" & UCase(Replace(txtCourseName, "'", "''")) & "'"
    Data4.RecordSource = SQLQ
    Data4.Refresh
    If Data4.Recordset.EOF Then
        MsgBox "Old Course Name in active employee's records does not exist", vbCritical, "Course Name Error"
        txtCourseName.SetFocus
        Exit Function
    End If

    chkFUComment = True
    Exit Function
End If

If Len(clpOldCode) = 0 Then
    MsgBox "Old code is a required field"
    clpOldCode.SetFocus
    Exit Function
End If
If chkSep = 0 Or Not glbLinamar Then
    If Len(clpNewCode) = 0 And fglbState = "M" Then
        MsgBox "New code is a required field"
        clpNewCode.SetFocus
        Exit Function
    End If
    
    If Len(clpNewCode) > 0 And clpNewCode.Caption = "Unassigned" Then
        MsgBox "If New Code Entered - it must be known"
        clpNewCode.SetFocus
        Exit Function
    End If
End If
If optCodes(1) And glbLinamar Then
    Dim rsEmp As New ADODB.Recordset
    rsEmp.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_DIV='" & clpNewCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        MsgBox "The New Facility has employees already."
        clpNewCode.SetFocus
        Exit Function
    End If
End If
If glbWFC Then
    'Ticket #12355
    If optCodes(0) Or optCodes(4) Or (optCodes(3) And txtTblName = "EDPT") Then
        If Len(clpCode(0).Text) = 0 Then
            MsgBox lStr("Section code is required")
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox lStr("If Section code entered it must be known")
        clpCode(0).SetFocus
        Exit Function
    End If
    If Len(clpCode(0).Text) = 0 Then
        Title$ = "Mass Code Master Remove"
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = lStr("The Section is blank. ") & Chr(10)
        Msg$ = Msg$ & lStr("Are You Sure You Want To Muss Update the code for all WFC Sections?")
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            clpCode(0).SetFocus
            Exit Function
        End If
    End If
End If

chkFUComment = True

Exit Function

chkFUComment_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkFUComment", "HR Attendance", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub chkSep_Click()
If Not chkSep.Visible Then Exit Sub
If chkSep = 0 Then
    cmbNewDIV.Visible = True
    clpNewCode.Visible = True
Else
    cmbNewDIV.Visible = False
    clpNewCode.Visible = False
End If
End Sub

Private Sub clpDiv_LostFocus()
    ' danielk - 12/30/2002 - moved all code from _Change to _LostFocus, Jaddy asked me to
    cmbNewDIV.Top = 5145
    clpNewCode.Top = 5145
    chkSep.Visible = False
    If Not glbLinamar Then Exit Sub
                
    Select Case Data1.Recordset!TD_NAME
    Case "EDSE", "EDRG", "BNCD", "HMOP", "HMLN"
        If clpDIV <> otxtDIV And clpDIV.Caption <> "Unassigned" Then
            cmbNewDIV.Clear
            If clpDIV = "" Then Exit Sub
            cmbNewDIV.AddItem "ALL"
            cmbNewDIV.ListIndex = 0
            If clpDIV = "ALL" Then
                chkSep.Visible = True
                cmbNewDIV.Visible = False
                clpNewCode.Visible = False
                cmbNewDIV.Top = 5490
                clpNewCode.Top = 5490
            Else
                chkSep.Visible = False
                cmbNewDIV.AddItem clpDIV
            End If
            'Call clpOldCode.RefreshDescription
        End If
    End Select
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Dim rsTF As New ADODB.Recordset 'Frank 4/25/2000
Dim CntRec


Title$ = "Mass Code Master Remove"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Remove ALL codes for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If


fglbState = "D"
If Not chkFUComment() Then Exit Sub

x% = modDelRecs()

Screen.MousePointer = DEFAULT
MsgBox "Records Deleted Successfully"

Exit Sub


Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Code Master", "Delete")
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

Public Sub cmdModify_Click()

Dim SQLQ As String
Dim Title$, Msg$, DgDef As Variant, Response%

'''On Error GoTo Mod_Err
fglbState = "M"

If Not chkFUComment() Then Exit Sub

Title$ = "Mass Update Code Master"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to update all Records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

If Not modUpdRecs() Then Exit Sub

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MsgBox "Records Updated Successfully"

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

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUCODE"
End Sub

Private Sub Form_Load()

glbOnTop = "FRMUCODE"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim x%

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
Data2.ConnectionString = glbAdoIHRDB
Data3.ConnectionString = glbAdoIHRDB
Data4.ConnectionString = glbAdoIHRDB
Data5.ConnectionString = glbAdoIHRAUDIT

Call setCaption(optCodes(0))
Call setCaption(optCodes(1))
Call setCaption(optCodes(2))
Call setCaption(optCodes(5))
Call setCaption(optCodes(6))
Call setCaption(optCodes(7))    'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name

lblSection.Caption = lStr(lblSection)

If glbWFC Then
    lblSection.Visible = True
    clpCode(0).Visible = True
    lblSection.Top = 5520
    clpCode(0).Top = 5520
End If

lblVadim.Visible = glbVadim

cmbNewDIV.Top = 5145
clpNewCode.Top = 5145

'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name - Opening up this function. It updates the Course Code
'for the matching Course Name.
'Ticket #22682 - Release 8.0. Jerry asked to comment this out as no one is using this function except Chatham-Kent but
'they are ORACLE client and we have stopped the upgrade for ORACLE since 7.8.
'optCodes(6).Visible = False

Call Retrieve

If optCodes(3) Then
    clpOldCode.MaxLength = 10 '8 '4 'Ticket #23166 Franks 01/29/2013
    clpNewCode.MaxLength = 10 '8 '4 'Ticket #23166 Franks 01/29/2013
    
    If txtTblName = "DOCT" Then     'Ticket #26353
        clpOldCode.MaxLength = 4
        clpNewCode.MaxLength = 4
    ElseIf txtTblName = "ESCD" Then  'Ticket #28852
        clpOldCode.MaxLength = 8
        clpNewCode.MaxLength = 8
    Else
        If txtTblName <> "EDAB" Then    'Ticket #30366 - Admin By allows 10 chrs
            clpOldCode.MaxLength = 4
            clpNewCode.MaxLength = 4
        End If
    End If
End If
Screen.MousePointer = DEFAULT

End Sub

Sub Retrieve()
Dim SQLQ, xTblName As String
Dim Tabl_Snap As New ADODB.Recordset
Dim x
SQLQ = "SELECT * FROM HRTABDES WHERE TD_NAME IN (SELECT CODENAME FROM HR_CODERELATE) ORDER BY TD_NAME"
Data1.RecordSource = SQLQ
Data1.Refresh
Do Until Data1.Recordset.EOF
    Data1.Recordset!TD_SHOWDESC = UCase(lStr(Data1.Recordset!TD_DESC))
    
    If glbCompSerial = "S/N - 2355W" Then
        If UCase(Data1.Recordset!TD_DESC) = "HIRE CODES" Then
            Data1.Recordset!TD_SHOWDESC = "VADIM STATUS CODE"
        End If
    End If
    
    Data1.Recordset.Update
    Data1.Recordset.MoveNext
Loop
If Data1.Recordset.RecordCount > 0 Then Data1.Recordset.MoveFirst
If glbLinamar Then
    Data3.RecordSource = "SELECT * FROM HR_DIVISION ORDER BY DIV"
    Data3.Refresh
    
    RowMax = Data3.Recordset.RecordCount
    ReDim ApptFacility(3, RowMax)
    x = 1
    Do Until Data3.Recordset.EOF
        If Data3.Recordset!Div <> "ALL" Then
            ApptFacility(1, x) = Data3.Recordset!Div
            x = x + 1
        End If
        Data3.Recordset.MoveNext
    Loop

End If

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmUCode = Nothing   'carmen apr 2000
End Sub

Private Function modDelRecs()
Dim SQLQ As String

modDelRecs = False
'''On Error GoTo modDelRecs_Err

Screen.MousePointer = HOURGLASS
If optCodes(1) And Not glbLinamar Then
    If Not modAUDIT("DIV", "D") Then Exit Function
    
    SQLQ = "UPDATE HREMP SET ED_DIV='' WHERE ED_DIV='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_HREMP SET ED_DIV='' WHERE ED_DIV='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HREMPHIS SET EE_NEWDIV='' WHERE EE_NEWDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HREMPHIS SET EE_OLDDIV='' WHERE EE_OLDDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMPHIS SET EE_NEWDIV='' WHERE EE_NEWDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn & ")"
    gdbAdoIhr001X.Execute SQLQ
    SQLQ = "UPDATE Term_HREMPHIS SET EE_OLDDIV='' WHERE EE_OLDDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HRVACENT SET VE_DIV='' WHERE VE_DIV='" & clpOldCode & "'"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HRPASDEP SET PD_DIV='' WHERE PD_DIV='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
ElseIf optCodes(2) Then
    If Not modAUDIT("GLNO", "D") Then Exit Function
    SQLQ = "UPDATE HREMP SET ED_GLNO='' WHERE ED_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_HREMP SET ED_GLNO='' WHERE ED_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
    'Ticket #16281 - Begin
    SQLQ = "UPDATE HRGLDIST SET GL_GLNO='' WHERE GL_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_HRGLDIST SET GL_GLNO='' WHERE GL_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_GLNO='' WHERE JH_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_GLNO='' WHERE JH_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
    SQLQ = "UPDATE HR_ATTENDANCE SET AD_GLNO='' WHERE AD_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_ATTENDANCE SET AD_GLNO='' WHERE AD_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
    SQLQ = "UPDATE HR_ATTENDANCE_HISTORY SET AH_GLNO='' WHERE AH_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRAUDIT SET AU_DEPT_GL='' WHERE AU_DEPT_GL='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
    'Ticket #16281 - End
ElseIf optCodes(5) Then
    If Not modAUDIT("SALDIST", "D") Then Exit Function
    SQLQ = "UPDATE HREMP SET ED_SALDIST='' WHERE ED_SALDIST='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE Term_HREMP SET ED_SALDIST='' WHERE ED_SALDIST='" & clpOldCode & "' AND " & glbSeleDeptUn
    gdbAdoIhr001X.Execute SQLQ
ElseIf optCodes(3) Then
    Data2.Refresh
    Dim oCodeName
    oCodeName = ""
    Dim xTLT As String
    Do Until Data2.Recordset.EOF
        If Data2.Recordset!CodeName <> oCodeName Then
            If Not modAUDIT(Data2.Recordset!CodeName & "", "D") Then Exit Function
        End If
        oCodeName = Data2.Recordset!CodeName
        If Data2.Recordset!TableName <> "HRAUDIT" Then
            SQLQ = "UPDATE " & Data2.Recordset!TableName
            SQLQ = SQLQ & " SET " & Data2.Recordset!FieldName & "='' "
            SQLQ = SQLQ & " WHERE " & Data2.Recordset!FieldName & "='" & IIf(clpDIV.Visible, clpDIV, "") & clpOldCode & "'"
            xTLT = Left(Data2.Recordset!FieldName, 3)
            If UCase(Left(Data2.Recordset!TableName, 4)) = "TERM" Then
                If xTLT <> "PP_" Then
                    If xTLT = "ED_" Then
                        SQLQ = SQLQ & " AND " & glbSeleDeptUn
                    Else
                        SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & "EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn & ")"
                    End If
                End If
                gdbAdoIhr001X.Execute SQLQ
            Else
                If xTLT <> "PP_" Then
                    If xTLT = "ED_" Then
                        SQLQ = SQLQ & " AND " & glbSeleDeptUn
                    Else
                        SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & "EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"
                    End If
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
        End If
        Data2.Recordset.MoveNext
    Loop
    Call tblDelRecs
End If

modDelRecs = True

Exit Function

modDelRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "Delete Codes", "Delete")
modDelRecs = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function modUpdRecs()
Dim SQLQ As String
Dim xEmpField
Dim xEMPNBRNAME As String
Dim xStr As String
Dim sql As String

modUpdRecs = False

'''On Error GoTo modUpdRecs_Err

Screen.MousePointer = HOURGLASS

If optCodes(3) Then
    Data2.Refresh
    Dim oCodeName
    oCodeName = ""
    Dim xTLT As String
    Do Until Data2.Recordset.EOF
        If Data2.Recordset!CodeName <> oCodeName Then
            If Not modAUDIT(Data2.Recordset!CodeName & "", "M") Then Exit Function
        End If
        oCodeName = Data2.Recordset!CodeName
        If Data2.Recordset!TableName <> "HRAUDIT" Then
            SQLQ = "UPDATE " & Data2.Recordset!TableName
            SQLQ = SQLQ & " SET " & Data2.Recordset!FieldName
            If glbLinamar And chkSep <> 0 And clpDIV = "ALL" Then
                xEmpField = Left(Data2.Recordset!FieldName, 3) + "EMPNBR"
                SQLQ = SQLQ & "=RIGHT(" & xEmpField & ",3)+'" & clpOldCode & "'"
            ElseIf (Data2.Recordset!TableName = "HR_JOB_HISTORY" Or Data2.Recordset!TableName = "Term_JOB_HISTORY") And Data2.Recordset!FieldName = "JH_REGION" Then
                SQLQ = SQLQ & "='" & clpNewCode & "'"
            Else
                SQLQ = SQLQ & "='" & IIf(cmbNewDIV.Visible, cmbNewDIV, "") & clpNewCode & "'"
            End If
            SQLQ = SQLQ & " WHERE " & Data2.Recordset!FieldName & "='" & IIf(clpDIV.Visible, clpDIV, "") & clpOldCode & "'"
            xTLT = Left(Data2.Recordset!FieldName, 3)
            
            xEMPNBRNAME = IIf(xTLT = "CR_" Or xTLT = "CT_" Or xTLT = "RC_", "Empnbr", "EMPNBR")

            
            
            If UCase(Left(Data2.Recordset!TableName, 4)) = "TERM" Then
                If xTLT = "ED_" And Not Data2.Recordset!DOCTABLE Then   '8.0 - Ticket #22682 - Not updating _DOC database table
                    SQLQ = SQLQ & " AND " & glbSeleDeptUn
                    If glbWFC And Len(clpCode(0).Text) > 0 Then
                        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
                    End If
                Else
                    If Data2.Recordset!TableName = "Term_HRTRMEMP" Then 'Hemu 06/17/2003 - Begin - Ticket # 4332, In Term_HRTRMEMP table Employee# field name is different
                        '8.0 - Ticket #22682 - Added _DOCTYPE_TABL and _DOCTYPE in Attached Database
                        If Data2.Recordset!DOCTABLE Then
                            SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & "Employee_Number" & " IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".dbo." & "Term_HREMP WHERE " & glbSeleDeptUn & ")"  'Hemu 06/17/2003 - End
                        Else
                            SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & "Employee_Number" & " IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn & ")"  'Hemu 06/17/2003 - End
                        End If
                    Else
                        '8.0 - Ticket #22682 - Added _DOCTYPE_TABL and _DOCTYPE in Attached Database
                        If Data2.Recordset!DOCTABLE Then
                            SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & xEMPNBRNAME & " IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".dbo." & "Term_HREMP WHERE " & glbSeleDeptUn '& ")"
                        Else
                            SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & xEMPNBRNAME & " IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn '& ")"
                        End If
                        If glbWFC And Len(clpCode(0).Text) > 0 Then
                            SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
                        End If
                        SQLQ = SQLQ & ")"
                    End If
                End If
                
                '8.0 - Ticket #22682 - Added _DOCTYPE_TABL and _DOCTYPE in Attached Database
                If Data2.Recordset!DOCTABLE Then
                    gdbAdoIhr001_DOC.Execute SQLQ
                Else
                    gdbAdoIhr001X.Execute SQLQ
                End If
            Else
                If xTLT = "ED_" And Not Data2.Recordset!DOCTABLE Then   '8.0 - Ticket #22682 - Not updating _DOC database table
                    SQLQ = SQLQ & " AND " & glbSeleDeptUn
                    If glbWFC And Len(clpCode(0).Text) > 0 Then
                        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
                    End If
                Else
                    If Data2.Recordset!TableName = "HRJOB" Or _
                        Data2.Recordset!TableName = "HRPASDEP" Or _
                        Data2.Recordset!TableName = "HRVACENT" Or _
                        Data2.Recordset!TableName = "HR_PAYPERIOD" Or _
                        Data2.Recordset!TableName = "HRSICKENT" Or _
                        Data2.Recordset!TableName = "HR_JOB_COURSE" Or _
                        Data2.Recordset!TableName = "HRJOBEVL" Or _
                        Data2.Recordset!TableName = "HRJOBSKL" Or _
                        Data2.Recordset!TableName = "HRJOBBUD" Or _
                        Data2.Recordset!TableName = "HRJOB_APP_PROCESS" Or _
                        Data2.Recordset!TableName = "HRJOB_RESP" Or _
                        Data2.Recordset!TableName = "HRJOB_DUTIES" Or _
                        Data2.Recordset!TableName = "HR_HOURLYENT" Or _
                        Data2.Recordset!TableName = "HRJOB_GRADE" Or _
                        Data2.Recordset!TableName = "HR_BENEFITS_GROUP" Or _
                        Data2.Recordset!TableName = "HR_BENEFIT_COST" Or _
                        Data2.Recordset!TableName = "HR_BENEFITS_GROUP_MATRIX" Or _
                        Data2.Recordset!TableName = "HR_COURSECODE_MASTER" Or _
                        Data2.Recordset!TableName = "HRP_PENSION_MASTER" Or _
                        Data2.Recordset!TableName = "HR_PERF_JOBGRP" Or _
                        Data2.Recordset!TableName = "HRBUDGET" Or _
                        Data2.Recordset!TableName = "HRATT_MATRIX" Or _
                        Data2.Recordset!TableName = "HRLABEL" Or Data2.Recordset!TableName = "HRA_APPLFORM_WRKFLOW" Or Data2.Recordset!TableName = "HRA_LETTER_POSTYPE" Or _
                        Data2.Recordset!TableName = "HR_SECURE_COMMENTS" Or Data2.Recordset!TableName = "HR_SECURE_FOLLOW_UP" Or Data2.Recordset!TableName = "HR_SECURE_ATTENDANCE" Or Data2.Recordset!TableName = "HRDOC_JOB" Or Data2.Recordset!TableName = "HR_SECURE_DOCUMENT_TYPE" Or _
                        Data2.Recordset!TableName = "HRA_SECURE_REQUISITION" Or Data2.Recordset!TableName = "HRVACPCTENT" Or Data2.Recordset!TableName = "HRVACENTDAILY" Or Data2.Recordset!TableName = "HR_DAILYVACACCR" Or Data2.Recordset!TableName = "HR_DAILYACC_LOG" Then    'Release 8.1, Release 8.2 (Ticket #30508)

                        SQLQ = SQLQ
                    Else
                        'Ticket #18501
                        If UCase(Data2.Recordset!TableName) = "HR_COMP_PRIORITY" Then
                           ' SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & xEMPNBRNAME & " IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn '& ")"
                        Else
                            '8.0 - Ticket #22682 - Added _DOCTYPE_TABL and _DOCTYPE in Attached Database
                            If Data2.Recordset!DOCTABLE Then
                                SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & xEMPNBRNAME & " IN (SELECT ED_EMPNBR FROM " & SQLDatabaseName & ".dbo." & "HREMP WHERE " & glbSeleDeptUn '& ")"
                            Else
                                SQLQ = SQLQ & " AND " & IIf(Right(xTLT, 1) = "_", xTLT, "") & xEMPNBRNAME & " IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn '& ")"
                            End If
                        End If
                        If glbWFC And Len(clpCode(0).Text) > 0 Then
                            SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
                        End If
                        'Ticket #18501 Added the Timesheet tables
                        If UCase(Data2.Recordset!TableName) <> "HR_COMP_PRIORITY" Then
                            SQLQ = SQLQ & ")"
                        End If
                    End If
                End If
                
                'Table HR_JOB_COURSE, [Job + Course Code ] is primary key. If replace old code with new code directly
                'It might cause duplicate key error, see ticket# 7725
                If Data2.Recordset!TableName = "HR_JOB_COURSE" And oCodeName = "ESCD" And Data2.Recordset!FieldName = "PC_CRSCODE" Then
                    Call modUpdRecHR_JOB_COURSE(clpOldCode, clpNewCode)
                Else
                    '8.0 - Ticket #22682 - Added _DOCTYPE_TABL and _DOCTYPE in Attached Database
                    If Data2.Recordset!DOCTABLE Then
                        gdbAdoIhr001_DOC.Execute SQLQ
                    'Ticket #23118
                    ElseIf Data2.Recordset!TableName = "HR_SECURE_COMMENTS" Or _
                        Data2.Recordset!TableName = "HR_SECURE_FOLLOW_UP" Or _
                        Data2.Recordset!TableName = "HR_SECURE_ATTENDANCE" Or _
                        Data2.Recordset!TableName = "HR_SECURE_DOCUMENT_TYPE" Then  'Release 8.1
                        'Do not update these tables because it will duplicate the code. When a new code is created, it
                        'automatically adds a new code in these secure tables. And so whoever created gets full
                        'permission on that code under these secure tables.
                        'At this point since this is a mass update function changing from one code to another, Delete
                        'the Older code from that SECURE table as the code itself will be deleted from the Table Master
                        'anyways in the later part of this code, i.e. '*-* below
                        gdbAdoIhr001.Execute "DELETE FROM " & Data2.Recordset!TableName & " WHERE " & Data2.Recordset!FieldName & "='" & IIf(clpDIV.Visible, clpDIV, "") & clpOldCode & "'"
                    ElseIf Data2.Recordset!TableName = "HR_COURSECODE_MASTER" And oCodeName = "ESCD" And Data2.Recordset!FieldName = "ES_CRSCODE" Then
                        'Ticket #23166 Franks 01/31/2013
                        gdbAdoIhr001.Execute "DELETE FROM " & Data2.Recordset!TableName & " WHERE " & Data2.Recordset!FieldName & "='" & clpOldCode & "'"
                    Else
                        gdbAdoIhr001.Execute SQLQ
                    End If
                End If
            End If
        End If
        Data2.Recordset.MoveNext
    Loop
    
    'Ticket #20484 Franks 06/17/2011
    'If glbWFC Then
        'Update Payroll Matrix
        SQLQ = "UPDATE HRMATRIX SET M_CODE = '" & clpNewCode & "' "
        SQLQ = SQLQ & "WHERE M_TYPE = '" & oCodeName & "' AND M_CODE = '" & clpOldCode & "' "
        If Len(clpCode(0).Text) > 0 Then
            SQLQ = SQLQ & "AND M_SECTION = '" & clpCode(0).Text & "' "
        End If
        gdbAdoIhr001.Execute SQLQ
        'Update Code Matrix
        SQLQ = "UPDATE CODEMATRIX SET CM_KEY = '" & clpNewCode & "' "
        SQLQ = SQLQ & "WHERE CM_NAME = '" & oCodeName & "' AND CM_KEY = '" & clpOldCode & "' "
        If Len(clpCode(0).Text) > 0 Then
            SQLQ = SQLQ & "AND CM_SECTION = '" & clpCode(0).Text & "' "
        End If
        gdbAdoIhr001.Execute SQLQ
    'End If
    
    'Ticket #28118 Franks 02/01/2016 - begin
    If glbWFC Then
        If oCodeName = "EDRG" Then 'for Region Code update
            If Len(clpOldCode) > 0 And Len(clpNewCode) > 0 Then
                SQLQ = "UPDATE HRJOB SET JB_REGION = '" & clpNewCode & "' "
                SQLQ = SQLQ & "WHERE JB_REGION = '" & clpOldCode & "' "
                If Len(clpCode(0).Text) > 0 Then
                    SQLQ = SQLQ & "AND M_SECTION = '" & clpCode(0).Text & "' "
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
        End If
    End If
    'Ticket #28118 Franks 02/01/2016 - end
    
    If glbVadim Then
        'Attendance, Benefit and Benefit Group Codes only
        If oCodeName = "ADRE" Then
            'VADIM_PAYCODE table
            xStr = "UPDATE VADIM_PAYCODE SET IHR_CODE = '" & clpNewCode & "'"
            xStr = xStr & " WHERE IHR_CODE = '" & clpOldCode & "'"
            gdbAdoIhr001.Execute xStr
            
            'VADIM_ACCRUAL_CLASS table
            xStr = "UPDATE VADIM_ACCRUAL_CLASS SET IHR_CODE = '" & clpNewCode & "'"
            xStr = xStr & " WHERE IHR_CODE = '" & clpOldCode & "'"
            gdbAdoIhr001.Execute xStr
                        
        ElseIf oCodeName = "BNCD" Then
            'HRMATRIX_BENEFIT table
            xStr = "UPDATE HRMATRIX_BENEFIT SET IHR_BCODE = '" & clpNewCode & "'"
            xStr = xStr & " WHERE IHR_BCODE = '" & clpOldCode & "'"
            gdbAdoIhr001.Execute xStr
        ElseIf oCodeName = "BGMF" Then
            'HRMATRIX_BENEFIT table
            xStr = "UPDATE HRMATRIX_BENEFIT SET IHR_BGROUP = '" & clpNewCode & "'"
            xStr = xStr & " WHERE IHR_BGROUP = '" & clpOldCode & "'"
            gdbAdoIhr001.Execute xStr
        End If
        
        'Update Code Matrix
        SQLQ = "UPDATE CODEMATRIX SET CM_KEY = '" & clpNewCode & "' "
        SQLQ = SQLQ & "WHERE CM_NAME = '" & oCodeName & "' AND CM_KEY = '" & clpOldCode & "' "
        gdbAdoIhr001.Execute SQLQ
    End If
    
    'Frank 03/26/04 Ticket# 5733
    'If WFC and Section entered, can't delete this Code, maybe other plant use it
    If (glbWFC And Len(clpCode(0).Text) > 0) Then
    Else
        '*-*
        Call tblDelRecs
    End If
    
    Call tblADDRecs
    
    If txtTblName = "EDPT" Then
        If Not glbCompSerial = "S/N - 2394W" Then  'St. John #14796
            Call Employee_Mass_Update_Integration(clpOldCode, clpNewCode, clpCode(0), "Employee Category Mass Update")
        End If
    End If
    
ElseIf optCodes(0) Then
    If Not glbCompSerial = "S/N - 2394W" Then  'St. John #14796
        Call Employee_Mass_Update_Integration(clpOldCode, clpNewCode, clpCode(0), "Employee Dept Mass Update")
    End If
    
    If Not modAUDIT("DEPT", "M") Then Exit Function

    SQLQ = "UPDATE HREMP SET ED_DEPTNO='" & clpNewCode & "' WHERE ED_DEPTNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMP SET ED_DEPTNO='" & clpNewCode & "' WHERE ED_DEPTNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HREMPHIS SET EE_OLDDEPT='" & clpNewCode & "' WHERE EE_OLDDEPT='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HREMPHIS SET EE_NEWDEPT='" & clpNewCode & "' WHERE EE_NEWDEPT='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMPHIS SET EE_OLDDEPT='" & clpNewCode & "' WHERE EE_OLDDEPT='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMPHIS SET EE_NEWDEPT='" & clpNewCode & "' WHERE EE_NEWDEPT='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HRVACENT SET VE_DEPT='" & clpNewCode & "' WHERE VE_DEPT='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRPASDEP SET PD_DEPT='" & clpNewCode & "' WHERE PD_DEPT='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_DEPTNO='" & clpNewCode & "' WHERE JH_DEPTNO='" & clpOldCode & "' AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_DEPTNO='" & clpNewCode & "' WHERE JH_DEPTNO='" & clpOldCode & "' AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HRVACPCTENT SET VP_DEPT='" & clpNewCode & "' WHERE VP_DEPT='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
ElseIf optCodes(1) Then
    If Not modAUDIT("DIV", "M") Then Exit Function
    
    SQLQ = "UPDATE HREMP SET ED_DIV='" & clpNewCode & "' WHERE ED_DIV='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMP SET ED_DIV='" & clpNewCode & "' WHERE ED_DIV='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    
    SQLQ = "UPDATE HREMPHIS SET EE_NEWDIV='" & clpNewCode & "' WHERE EE_NEWDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HREMPHIS SET EE_OLDDIV='" & clpNewCode & "' WHERE EE_OLDDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMPHIS SET EE_NEWDIV='" & clpNewCode & "' WHERE EE_NEWDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMPHIS SET EE_OLDDIV='" & clpNewCode & "' WHERE EE_OLDDIV='" & clpOldCode & "' AND EE_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn '& ")"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HRVACENT SET VE_DIV='" & clpNewCode & "' WHERE VE_DIV='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRPASDEP SET PD_DIV='" & clpNewCode & "' WHERE PD_DIV='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
   
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_DIV='" & clpNewCode & "' WHERE JH_DIV='" & clpOldCode & "' AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_DIV='" & clpNewCode & "' WHERE JH_DIV='" & clpOldCode & "' AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
   
    SQLQ = "UPDATE HRVACPCTENT SET VP_DIV='" & clpNewCode & "' WHERE VP_DIV='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
   
    If glbLinamar Then
        SQLQ = "UPDATE HRAUDIT SET AU_TIDIV='" & clpNewCode & "' WHERE AU_TIDIV='" & clpOldCode & "'"
        gdbAdoIhr001X.Execute SQLQ
        SQLQ = "UPDATE HRAUDIT SET AU_TODIV='" & clpNewCode & "' WHERE AU_TODIV='" & clpOldCode & "'"
        gdbAdoIhr001X.Execute SQLQ
        Call BatUpdateDIV
    End If
    
ElseIf optCodes(2) Then
    If Not modAUDIT("GLNO", "M") Then Exit Function
    
    SQLQ = "UPDATE HREMP SET ED_GLNO='" & clpNewCode & "' WHERE ED_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMP SET ED_GLNO='" & clpNewCode & "' WHERE ED_GLNO='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    'Ticket #16281 - Begin
    SQLQ = "UPDATE HRGLDIST SET GL_GLNO='" & clpNewCode & "' WHERE GL_GLNO='" & clpOldCode & "' AND GL_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND GL_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HRGLDIST SET GL_GLNO='" & clpNewCode & "' WHERE GL_GLNO='" & clpOldCode & "' AND GL_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND GL_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE  ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_GLNO='" & clpNewCode & "' WHERE JH_GLNO='" & clpOldCode & "' AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_GLNO='" & clpNewCode & "' WHERE JH_GLNO='" & clpOldCode & "' AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE  ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HR_ATTENDANCE SET AD_GLNO='" & clpNewCode & "' WHERE AD_GLNO='" & clpOldCode & "' AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_ATTENDANCE SET AD_GLNO='" & clpNewCode & "' WHERE AD_GLNO='" & clpOldCode & "' AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE  ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HR_ATTENDANCE_HISTORY SET AH_GLNO='" & clpNewCode & "' WHERE AH_GLNO='" & clpOldCode & "' AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRAUDIT SET AU_DEPT_GL='" & clpNewCode & "' WHERE AU_DEPT_GL='" & clpOldCode & "' AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        'SQLQ = SQLQ & " AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE  ED_SECTION = '" & clpCode(0).Text & "') "
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    SQLQ = SQLQ & ")"
    gdbAdoIhr001X.Execute SQLQ
    'Ticket #16281 - End
    
ElseIf optCodes(5) Then
    If Not modAUDIT("SALDIST", "M") Then Exit Function
    
    SQLQ = "UPDATE HREMP SET ED_SALDIST='" & clpNewCode & "' WHERE ED_SALDIST='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREMP SET ED_SALDIST='" & clpNewCode & "' WHERE ED_SALDIST='" & clpOldCode & "' AND " & glbSeleDeptUn
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
ElseIf optCodes(4) Then
    If Not glbCompSerial = "S/N - 2394W" Then  'St. John #14796
        Call Employee_Mass_Update_Integration(clpOldCode, clpNewCode, clpCode(0), "Employee JobCode Mass Update")
    End If
    
    If Not modAUDIT("JOB", "M") Then Exit Function
'
'    SQLQ = "UPDATE HRJOB SET JB_CODE='" & clpNewCode & "' WHERE JB_CODE='" & clpOldCode & "'"
'    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HR_JOB_HISTORY SET JH_JOB='" & clpNewCode & "' WHERE JH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_JOB='" & clpNewCode & "' WHERE SH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_PERFORM_HISTORY SET PH_JOB='" & clpNewCode & "' WHERE PH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND PH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_ATTENDANCE SET AD_JOB='" & clpNewCode & "' WHERE AD_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_ATTENDANCE_HISTORY SET AH_JOB='" & clpNewCode & "' WHERE AH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_JBCODE='" & clpNewCode & "' WHERE EC_JBCODE='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND EC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOBSKL SET JS_CODE='" & clpNewCode & "' WHERE JS_CODE='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOBEVL SET JE_CODE='" & clpNewCode & "' WHERE JE_CODE='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOBBUD SET JG_CODE='" & clpNewCode & "' WHERE JG_CODE='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOB_APP_PROCESS SET JA_JOB='" & clpNewCode & "' WHERE JA_JOB='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
        
    SQLQ = "UPDATE HRJOB_RESP SET JR_JOB='" & clpNewCode & "' WHERE JR_JOB='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOB_DUTIES SET JD_JOB='" & clpNewCode & "' WHERE JD_JOB='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HRJOB_GRADE SET JB_CODE='" & clpNewCode & "' WHERE JB_CODE='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
        
    SQLQ = "UPDATE HR_JOB_COURSE SET PC_JOB='" & clpNewCode & "' WHERE PC_JOB='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
        
     If glbOttawaCCAC Then
         SQLQ = "UPDATE HR_JOB_CONTROL SET PC_JOB='" & clpNewCode & "' WHERE PC_JOB='" & clpOldCode & "'"
        gdbAdoIhr001.Execute SQLQ
    End If

    SQLQ = "UPDATE HRJOB SET JB_REPTAU='" & clpNewCode & "' WHERE JB_REPTAU='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRJOB SET JB_REPTAU2='" & clpNewCode & "' WHERE JB_REPTAU2='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRJOB SET JB_REPTAU3='" & clpNewCode & "' WHERE JB_REPTAU3='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ

    SQLQ = "UPDATE Term_JOB_HISTORY SET JH_JOB='" & clpNewCode & "' WHERE JH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE Term_SALARY_HISTORY SET SH_JOB='" & clpNewCode & "' WHERE SH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE Term_PERFORM_HISTORY SET PH_JOB='" & clpNewCode & "' WHERE PH_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND pH_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE Term_ATTENDANCE SET AD_JOB='" & clpNewCode & "' WHERE AD_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE Term_HR_OCC_HEALTH_SAFETY SET EC_JBCODE='" & clpNewCode & "' WHERE EC_JBCODE='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND EC_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    'Friesens - Ticket #16189
    SQLQ = "UPDATE HR_TEMP_WORK SET TW_JOB='" & clpNewCode & "' WHERE TW_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND TW_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE TERM_TEMP_WORK SET TW_JOB='" & clpNewCode & "' WHERE TW_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND TW_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HR_JOB_DOCUMENT SET JD_JOB='" & clpNewCode & "' WHERE JD_JOB='" & clpOldCode & "'"
    'If glbWFC And Len(clpCode(0).Text) > 0 Then
    '    SQLQ = SQLQ & " AND JD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    'End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE HREDSEM SET ES_JOB='" & clpNewCode & "' WHERE ES_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_HREDSEM SET ES_JOB='" & clpNewCode & "' WHERE ES_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND ES_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    SQLQ = "UPDATE HR_TRAIN SET TR_JOB='" & clpNewCode & "' WHERE TR_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND TR_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    'Ticket #24410 - City of Sarnia - new job field
    SQLQ = "UPDATE HREARN SET OE_JOB='" & clpNewCode & "' WHERE OE_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "UPDATE Term_EARN SET OE_JOB='" & clpNewCode & "' WHERE OE_JOB='" & clpOldCode & "'"
    If glbWFC And Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & " AND EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
    End If
    gdbAdoIhr001X.Execute SQLQ
    
    'Ticket #30508 - Applicant Tracking Enhancement
    SQLQ = "UPDATE HRA_APPLFORM_WRKFLOW SET WF_AUTH_JOB1='" & clpNewCode & "' WHERE WF_AUTH_JOB1='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRA_APPLFORM_WRKFLOW SET WF_AUTH_JOB2='" & clpNewCode & "' WHERE WF_AUTH_JOB2='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRA_APPLFORM_WRKFLOW SET WF_AUTH_JOB3='" & clpNewCode & "' WHERE WF_AUTH_JOB3='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRA_APPLFORM_WRKFLOW SET WF_AUTH_JOB4='" & clpNewCode & "' WHERE WF_AUTH_JOB4='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRA_APPLFORM_WRKFLOW SET WF_AUTH_JOB5='" & clpNewCode & "' WHERE WF_AUTH_JOB5='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
    'SQLQ = "UPDATE HRA_SECURE_REQUISITION SET RS_EXCLJOB='" & clpNewCode & "' WHERE RS_EXCLJOB='" & clpOldCode & "'"
    'gdbAdoIhr001.Execute SQLQ
    'SQLQ = "UPDATE HRA_SECURE_REQUISITION SET RS_INCLJOB='" & clpNewCode & "' WHERE RS_INCLJOB='" & clpOldCode & "'"
    'gdbAdoIhr001.Execute SQLQ
    
ElseIf optCodes(6) Then
   
    'Do While Not Data5.Recordset.EOF
        sql = "UPDATE Term_HREDSEM SET ES_CRSCODE='" & clpCode(1) & "'"
        'Ticket #23369
        'sql = sql & "WHERE ES_COURSE = '" & Replace(txtCourseName, "'", "'+chr(39)+'") & "'"
        sql = sql & " WHERE UPPER(ES_COURSE) = '" & UCase(Replace(txtCourseName, "'", "''")) & "'"
        gdbAdoIhr001X.Execute sql
        'Data5.Recordset.MoveNext
    'Loop
   
    'Do While Not Data4.Recordset.EOF
        sql = "UPDATE HREDSEM SET ES_CRSCODE='" & clpCode(1) & "'"
        'Ticket #23369
        'sql = sql & "WHERE ES_COURSE = '" & Replace(txtCourseName, "'", "'+chr(39)+'") & "'"
        sql = sql & " WHERE UPPER(ES_COURSE) = '" & UCase(Replace(txtCourseName, "'", "''")) & "'"
        gdbAdoIhr001.Execute sql
        'Data4.Recordset.MoveNext
    'Loop

ElseIf optCodes(7) Then
    'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
   
    'Not updating Term Employees
    'Do While Not Data5.Recordset.EOF
    '    sql = "UPDATE Term_HREDSEM SET ES_COURSE='" & Replace(txtNewCourseName, "'", "''") & "'"
    '    sql = sql & "WHERE ES_COURSE = '" & Replace(txtCourseName, "'", "''") & "'"
    '    gdbAdoIhr001X.Execute sql
    '    Data5.Recordset.MoveNext
    'Loop
   
    'Do While Not Data4.Recordset.EOF
        sql = "UPDATE HREDSEM SET ES_COURSE='" & Replace(txtNewCourseName, "'", "''") & "'"
        sql = sql & " WHERE UPPER(ES_COURSE) = '" & UCase(Replace(txtCourseName, "'", "''")) & "'"
        gdbAdoIhr001.Execute sql
    '    Data4.Recordset.MoveNext
    'Loop
    
End If

modUpdRecs = True

Exit Function

modUpdRecs_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modUpdRecs", "Update Codes", "Update")
modUpdRecs = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub modUpdRecHR_JOB_COURSE(xOldCode, xNewCode)
Dim rsTemp As New ADODB.Recordset
Dim rsTemp01 As New ADODB.Recordset
Dim SQLQ
    SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_CRSCODE='" & xOldCode & "' "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsTemp.EOF
        SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB='" & rsTemp("PC_JOB") & "' AND PC_CRSCODE='" & xNewCode & "' "
        If rsTemp01.State <> 0 Then rsTemp01.Close
        rsTemp01.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp01.EOF Then
            rsTemp.Delete
        Else
            rsTemp("PC_CRSCODE") = xNewCode
            rsTemp.Update
        End If
        rsTemp01.Close
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
End Sub

Private Sub optCodes_Click(Index As Integer)
lblMsg.Visible = False

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpOldCode.TransDiv = ""
    clpNewCode.TransDiv = ""
End If
        
If Index = 3 Then
    tblTables.Visible = True
    lblTitle(0).Visible = True
    txtTblName.Visible = True
    lblTblNameDesc.Visible = True
    clpOldCode.Visible = True
    clpNewCode.Visible = True
    lblOldCode.Visible = True
    lblNewCode.Visible = True
    'Ticket #26247 - if Benefit Code then allow 10 char
    'Ticket #30366 - Admin By allows 10 chrs
    If txtTblName = "BNCD" Or txtTblName = "EDAB" Then
        clpOldCode.MaxLength = 10
        clpNewCode.MaxLength = 10
    ElseIf txtTblName = "DOCT" Then     'Ticket #26353
        clpOldCode.MaxLength = 4
        clpNewCode.MaxLength = 4
    ElseIf txtTblName = "ESCD" Then  'Ticket #28852
        clpOldCode.MaxLength = 8
        clpNewCode.MaxLength = 8
    Else
        clpOldCode.MaxLength = 4    '10 '8 '4 'Ticket #23166 Franks 01/29/2013
        clpNewCode.MaxLength = 4    '10 '8 '4 'Ticket #23166 Franks 01/29/2013
    End If
    clpCode(1).Visible = False
    txtCourseName.Visible = False
    lblCourseName.Visible = False
    lblCourseCode.Visible = False
    
    Call setCodesLayout

    'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
    txtNewCourseName.Visible = False
    lblNewCourseName.Visible = False

Else
    If Index = 0 Then
        clpOldCode.LookupType = Department
        clpNewCode.LookupType = Department
        clpOldCode.MaxLength = 7
        clpNewCode.MaxLength = 7
        clpCode(1).Visible = False
        txtCourseName.Visible = False
        lblCourseName.Visible = False
        lblCourseCode.Visible = False
        clpOldCode.Visible = True
        clpNewCode.Visible = True
        lblOldCode.Visible = True
        lblNewCode.Visible = True
    
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        txtNewCourseName.Visible = False
        lblNewCourseName.Visible = False
    
    ElseIf Index = 1 Then
        clpOldCode.LookupType = Division
        clpNewCode.LookupType = Division
        clpOldCode.MaxLength = 4
        clpNewCode.MaxLength = 4
        clpCode(1).Visible = False
        txtCourseName.Visible = False
        lblCourseName.Visible = False
        lblCourseCode.Visible = False
        clpOldCode.Visible = True
        clpNewCode.Visible = True
        lblOldCode.Visible = True
        lblNewCode.Visible = True
    
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        txtNewCourseName.Visible = False
        lblNewCourseName.Visible = False
    
    ElseIf Index = 2 Then
        clpOldCode.LookupType = GL
        clpNewCode.LookupType = GL
        clpOldCode.MaxLength = 20
        clpNewCode.MaxLength = 20
        clpCode(1).Visible = False
        txtCourseName.Visible = False
        lblCourseName.Visible = False
        lblCourseCode.Visible = False
        clpOldCode.Visible = True
        clpNewCode.Visible = True
        lblOldCode.Visible = True
        lblNewCode.Visible = True
    
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        txtNewCourseName.Visible = False
        lblNewCourseName.Visible = False
    
    ElseIf Index = 4 Then
        clpOldCode.LookupType = Job
        clpNewCode.LookupType = Job
        If glbWFC Then 'Ticket #25911 Franks 10/21/2014
            clpOldCode.TransDiv = glbWFCUserSecList
            clpNewCode.TransDiv = glbWFCUserSecList
        End If
        clpOldCode.MaxLength = 25 ' 6
        clpNewCode.MaxLength = 25 ' 6
        clpCode(1).Visible = False
        txtCourseName.Visible = False
        lblCourseName.Visible = False
        lblCourseCode.Visible = False
        clpOldCode.Visible = True
        clpNewCode.Visible = True
        lblOldCode.Visible = True
        lblNewCode.Visible = True
    
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        txtNewCourseName.Visible = False
        lblNewCourseName.Visible = False
    
    ElseIf Index = 5 Then
        clpOldCode.LookupType = SalaryDistribution
        clpNewCode.LookupType = SalaryDistribution
        clpOldCode.MaxLength = 6
        clpNewCode.MaxLength = 6
        clpCode(1).Visible = False
        txtCourseName.Visible = False
        lblCourseName.Visible = False
        lblCourseCode.Visible = False
        
        clpOldCode.Visible = True
        clpNewCode.Visible = True
        lblOldCode.Visible = True
        lblNewCode.Visible = True
    
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        txtNewCourseName.Visible = False
        lblNewCourseName.Visible = False
    
    ElseIf Index = 6 Then
        clpCode(1).Visible = True
        txtCourseName.Visible = True
        lblCourseName.Caption = "Course Name"
        lblCourseName.Visible = True
        lblCourseCode.Visible = True
        clpOldCode.Visible = False
        clpNewCode.Visible = False
        lblOldCode.Visible = False
        lblNewCode.Visible = False
        
        clpCode(1).Top = 3000
        txtCourseName.Top = 2500
        lblCourseName.Top = 2500
        lblCourseCode.Top = 3000
        
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        txtNewCourseName.Visible = False
        lblNewCourseName.Visible = False
        
        'Ticket #22682 - Release 8.0 - Jerry asked to display message
        lblMsg.Visible = True
        lblMsg.Left = lblCourseCode.Left
        lblMsg.Top = lblCourseCode.Top + 600
            
    ElseIf Index = 7 Then
        'Ticket #22682 - Release 8.0 - Old Course Name -> New Course Name
        clpCode(1).Visible = False
        lblCourseCode.Visible = False
        
        txtCourseName.Visible = True
        lblCourseName.Caption = "Old Course Name"
        lblCourseName.Visible = True
                
        txtNewCourseName.Visible = True
        lblNewCourseName.Visible = True
                
        clpOldCode.Visible = False
        clpNewCode.Visible = False
        lblOldCode.Visible = False
        lblNewCode.Visible = False
        
        txtCourseName.Top = 2500
        lblCourseName.Top = 2537
        txtNewCourseName.Top = 3000
        lblNewCourseName.Top = 3037
    End If
    
    lblTitle(1).Visible = False
    lblTitle(2).Visible = False
    clpDIV = ""
    clpDIV.Visible = False
    cmbNewDIV.Visible = False
    clpOldCode.Left = 1800
    clpNewCode.Left = 1800
    clpOldCode.TextBoxWidth = 1500
    clpNewCode.TextBoxWidth = 1500
    tblTables.Visible = False
    lblTitle(0).Visible = False
    txtTblName.Visible = False
    lblTblNameDesc.Visible = False
    clpOldCode.Tag = "OLD " & optCodes(Index).Caption
    clpNewCode.Tag = "NEW " & optCodes(Index).Caption
    clpOldCode = ""
    clpNewCode = ""
 '   If optCodes(1) Or (optCodes(2) And glbLinamar) Then
    '    cmdDelete.Enabled = False
 '   Else
    '    cmdDelete.Enabled = True
 '   End If
    Call INI_Controls(Me)
    Call SET_UP_MODE
End If
End Sub

Private Sub tblTables_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If tblTables.Tag = "ASC" Then
            tblTables.Tag = "DESC"
        Else
            tblTables.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRTABDES WHERE TD_NAME IN (SELECT CODENAME FROM HR_CODERELATE)"
        SQLQ = SQLQ & " ORDER BY " & tblTables.Columns(ColIndex).DataField & " " & tblTables.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub tblTables_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call setCodesLayout
End Sub

Private Sub setCodesLayout()

If Data1.Recordset.EOF Then Exit Sub

Data2.RecordSource = "SELECT * FROM HR_CODERELATE WHERE CODENAME='" & Data1.Recordset!TD_NAME & "'"
Data2.Refresh
clpOldCode.Tag = Data1.Recordset!TD_NAME & "-OLD " & UCase(Data1.Recordset!TD_DESC)
clpNewCode.Tag = Data1.Recordset!TD_NAME & "-NEW " & UCase(Data1.Recordset!TD_DESC)

clpOldCode.LookupType = HRTABL
clpOldCode.TablName = Data1.Recordset!TD_NAME

clpNewCode.LookupType = HRTABL
clpNewCode.TablName = Data1.Recordset!TD_NAME

If glbCompSerial = "S/N - 2355W" Then
    If Data1.Recordset!TD_NAME = "EDHC" Then
        clpOldCode.TABLTitle = "VADIM STATUS CODE"
        clpNewCode.TABLTitle = "VADIM STATUS CODE"
    End If
End If

'Ticket #26247 - if Benefit Code then allow 10 char
'Ticket #30366 - Admin By allows 10 chrs
If txtTblName = "BNCD" Or txtTblName = "EDAB" Then
    clpOldCode.MaxLength = 10
    clpNewCode.MaxLength = 10
ElseIf txtTblName = "DOCT" Then     'Ticket #26353
    clpOldCode.MaxLength = 4
    clpNewCode.MaxLength = 4
ElseIf txtTblName = "ESCD" Then  'Ticket #28852
    clpOldCode.MaxLength = 8
    clpNewCode.MaxLength = 8
Else
    clpOldCode.MaxLength = 4
    clpNewCode.MaxLength = 4
End If

Call INI_Controls(Me)

clpOldCode = ""
clpNewCode = ""
clpDIV = ""
cmbNewDIV.ListIndex = -1


MDIMain.MainToolBar.ButtonS("massdelete").Enabled = IIf(Data2.Recordset!REMOVABLE <> 0, True, False)

Dim xFacilityCode As Boolean

xFacilityCode = False

If glbLinamar Then
    Select Case Data1.Recordset!TD_NAME
    Case "EDSE", "EDRG", "BNCD", "HMOP", "HMLN"
        xFacilityCode = True
    End Select
End If
If xFacilityCode Then
    lblTitle(1).Visible = True
    clpDIV.Visible = True
    cmbNewDIV.Visible = True
    clpOldCode.Left = 2800
    clpNewCode.Left = 2800
    clpOldCode.TextBoxWidth = 3500
    clpNewCode.TextBoxWidth = 3500
    lblOldCode = "Old Information"
    lblNewCode = "New Information"
    lblTitle(1).Visible = True
    lblTitle(2).Visible = True
Else
    lblTitle(1).Visible = False
    clpDIV = ""
    clpDIV.Visible = False
    cmbNewDIV.Visible = False
    clpOldCode.Left = 1800
    clpNewCode.Left = 1800
    clpOldCode.TextBoxWidth = 1500
    clpNewCode.TextBoxWidth = 1500
    lblOldCode = "Old Code"
    lblNewCode = "New Code"
    lblTitle(1).Visible = False
    lblTitle(2).Visible = False
End If
End Sub

Private Sub clpDIV_Change()
    ' danielk - 12/30/2002 - moved code to LostFocus as Jaddy asked me to
End Sub

Private Function tblDelRecs()
Dim SQLQ As String, xTableName
If Left(Data1.Recordset!TD_NAME, 2) = "HM" Then
    xTableName = "LN_HOMES"
Else
    xTableName = "HRTABL"
End If

SQLQ = "DELETE FROM " & xTableName
SQLQ = SQLQ & " WHERE TB_NAME='" & Data1.Recordset!TD_NAME & "' AND TB_KEY='" & clpDIV & clpOldCode & "'"
gdbAdoIhr001.Execute SQLQ
End Function

Private Function tblADDRecs()
Dim x
Dim SQLQ As String, xTableName
Dim rsTB As New ADODB.Recordset
If Left(Data1.Recordset!TD_NAME, 2) = "HM" Then
    xTableName = "LN_HOMES"
Else
    xTableName = "HRTABL"
End If
If glbLinamar Then
    If clpDIV = "ALL" And chkSep <> 0 Then
        For x = 1 To UBound(ApptFacility, 2)
            SQLQ = "SELECT * FROM " & xTableName
            SQLQ = SQLQ & " WHERE TB_NAME='" & Data1.Recordset!TD_NAME & "'"
            SQLQ = SQLQ & " AND TB_KEY='" & ApptFacility(1, x) & ApptFacility(2, x) & "'"
            rsTB.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
            If rsTB.EOF Then
                rsTB.AddNew
                rsTB("TB_COMPNO") = "001"
                rsTB("TB_NAME") = Data1.Recordset!TD_NAME
                rsTB("TB_KEY") = ApptFacility(1, x) & ApptFacility(2, x)
                rsTB("TB_DESC") = clpOldCode.Caption
                rsTB("TB_LUSER") = glbUserID
                rsTB("TB_LDATE") = Date
                rsTB("TB_LTIME") = Time$
                rsTB.Update
            End If
            rsTB.Close
        Next
    End If
End If
End Function

Private Sub LGR_Desc(CNTR As Control, lblDesc As Label)
Dim rsGL  As New ADODB.Recordset
Dim SQLQ As String

'''On Error GoTo Jobd_Err
lblDesc.Visible = False
lblDesc.Caption = "Unassigned"
If Len(CNTR.Text) > 0 Then
  SQLQ = "SELECT * FROM HRGL WHERE GL_NO = '" & CStr(CNTR.Text) & "'"
  rsGL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
  If Not rsGL.EOF Then
      lblDesc.Caption = rsGL("GL_DESCR")
  End If
  lblDesc.Visible = True
Else

End If

Exit Sub

Jobd_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "LEAGERS", "SELECT")
Call RollBack '21June99 js

End Sub

Private Function GetRelativeBookmark(Bookmark As Variant, Offset As Long) As Variant
    Dim Index As Long
    Index = IndexFromBookmark(Bookmark, Offset)
    If Index <= 0 Or Index > RowMax Then
        GetRelativeBookmark = Null
    Else
        GetRelativeBookmark = Str$(Index)
    End If
End Function

Private Function IndexFromBookmark(Bookmark As Variant, Offset As Long) As Long
    Dim Index As Long
    If IsNull(Bookmark) Then
        If Offset <= 0 Then
            Index = (RowMax + 1) + Offset
        Else
            Index = 0 + Offset
        End If
    Else
        Index = Val(Bookmark) + Offset
    End If
    If Index > 0 And Index < RowMax Then
       IndexFromBookmark = Index
    Else
       IndexFromBookmark = -9999
    End If
End Function

Private Function GetUserData(Bookmark As Variant, col As Integer) As Variant
    Dim Index As Long
        
    Index = IndexFromBookmark(Bookmark, 0)
    If Index <= 0 Or Index > RowMax Then
        GetUserData = Null
    Else
    GetUserData = ApptFacility(col + 1, Index)
    End If
End Function

Private Function StoreUserData(Bookmark As Variant, col As Integer, Userval As Variant) As Boolean
    Dim Index As Long
    Index = IndexFromBookmark(Bookmark, 0)
    If Index <= 0 Or Index > RowMax Then
        StoreUserData = False
    Else
        StoreUserData = True
        ApptFacility(col + 1, Index) = Userval
   End If
End Function

Private Function BatUpdateDIV()
Dim dyn_Table As New ADODB.Recordset
Dim xCount, xx
Dim SQLQ, x%, xFldTitle, xFld As String, xTable As String
BatUpdateDIV = False
'''On Error GoTo BatUpdateDIV_cmdUpdErr
Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0

SQLQ = "SELECT * FROM INFO_HR_TABLES "
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
SQLQ = SQLQ & " WHERE (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"

dyn_Table.Open SQLQ, gdbAdoIhr001, adOpenStatic
MDIMain.panHelp(0).FloodPercent = 10

xCount = dyn_Table.RecordCount
xx = 0
If Not modAUDIT("EMPNO", "M") Then Exit Function
Do Until dyn_Table.EOF
    MDIMain.panHelp(0).FloodPercent = (xx / xCount) * 60 + 10
    xTable = dyn_Table("Table_Name")
    
    If IsNull(dyn_Table("EMPNBR_Alias")) Then xFld = "" Else xFld = dyn_Table("EMPNBR_Alias")
    If InStr(xFld, "_") = 0 Then xFldTitle = "" Else xFldTitle = Left(xFld, 3)
    
    If dyn_Table("Employee_Keyed") Then
        Call UpdateEMPNBR(xTable, xFldTitle & "EMPNBR", xFldTitle)
        
        Select Case xTable
        Case "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY"
            Call UpdateEMPNBR(xTable, "SUPER", xFldTitle)
        Case "HR_JOB_HISTORY", "HR_PERFORM_HISTORY"
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU", xFldTitle)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU2", xFldTitle)
            Call UpdateEMPNBR(xTable, xFldTitle & "REPTAU3", xFldTitle)
        Case "HR_OCC_HEALTH_SAFETY"
            Call UpdateEMPNBR(xTable, xFldTitle & "EMPNOT", xFldTitle)
            Call UpdateEMPNBR(xTable, xFldTitle & "SUPERVISOR", xFldTitle)
        End Select
    End If
    dyn_Table.MoveNext
    xx = xx + 1
Loop
Call UpdateEMPNBR("Term_HRTRMEMP", "Employee_Number", "")
Call UpdateEMPNBR("HRAUDIT", "AU_EMPNBR", "AU_")
Call UpdateEMPNBR("HRAUDIT", "AU_TIEMPNBR", "AU_")
Call UpdateEMPNBR("HRAUDIT", "AU_TOEMPNBR", "AU_")

MDIMain.panHelp(0).FloodPercent = 70

Data2.RecordSource = "SELECT * FROM HR_CODERELATE WHERE CODENAME IN ('EDSE', 'EDRG', 'BNCD', 'HMOP', 'HMLN') ORDER BY CODENAME"
Data2.Refresh


Dim oCodeName
oCodeName = ""
Do Until Data2.Recordset.EOF
    If Data2.Recordset!CodeName <> oCodeName Then
        If Not modAUDIT(Data2.Recordset!CodeName & "", "M") Then Exit Function
    End If
    oCodeName = Data2.Recordset!CodeName
    If Data2.Recordset!TableName <> "HRAUDIT" Then
        SQLQ = "UPDATE " & Data2.Recordset!TableName
        SQLQ = SQLQ & " SET " & Data2.Recordset!FieldName & "='" & clpNewCode & "'+SUBSTRING(" & Data2.Recordset!FieldName & ",4,50)"
        SQLQ = SQLQ & " WHERE LEFT(" & Data2.Recordset!FieldName & ",3)='" & clpOldCode & "'"
        If UCase(Left(Data2.Recordset!TableName, 4)) = "TERM" Then
            gdbAdoIhr001X.Execute SQLQ
        Else
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    Data2.Recordset.MoveNext
Loop
MDIMain.panHelp(0).FloodPercent = 80
Dim rsTB As New ADODB.Recordset
SQLQ = "SELECT * FROM HRTABL"
SQLQ = SQLQ & " WHERE TB_NAME IN ('EDSE', 'EDRG', 'BNCD')"
SQLQ = SQLQ & " AND TB_KEY='" & clpNewCode & "'+SUBSTRING(TB_KEY,4,50)"
rsTB.Open SQLQ, gdbAdoIhr001, adOpenDynamic
If rsTB.EOF Then
    SQLQ = "UPDATE HRTABL "
    SQLQ = SQLQ & " SET TB_KEY='" & clpNewCode & "'+SUBSTRING(TB_KEY,4,50)"
    SQLQ = SQLQ & " WHERE TB_NAME IN ('EDSE', 'EDRG', 'BNCD') AND LEFT(TB_KEY,3)='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
End If
rsTB.Close
SQLQ = "SELECT * FROM LN_HOMES"
SQLQ = SQLQ & " WHERE TB_NAME IN ('EDSE', 'EDRG', 'BNCD')"
SQLQ = SQLQ & " AND TB_KEY='" & clpNewCode & "'+SUBSTRING(TB_KEY,4,50)"
rsTB.Open SQLQ, gdbAdoIhr001, adOpenDynamic
If rsTB.EOF Then
    SQLQ = "UPDATE LN_HOMES "
    SQLQ = SQLQ & " SET TB_KEY='" & clpNewCode & "'+SUBSTRING(TB_KEY,4,50)"
    SQLQ = SQLQ & " WHERE TB_NAME IN ( 'HMOP', 'HMLN') AND LEFT(TB_KEY,3)='" & clpOldCode & "'"
    gdbAdoIhr001.Execute SQLQ
End If
MDIMain.panHelp(0).FloodPercent = 100
Screen.MousePointer = DEFAULT
BatUpdateDIV = True

Exit Function

BatUpdateDIV_cmdUpdErr:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MDIMain.panHelp(0).FloodType = 0

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "BatUpdateDIV Error", xTable, "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub UpdateEMPNBR(nTable As String, nFld As String, nFldTitle)
Dim SQLQ
SQLQ = "UPDATE " & nTable & " SET "
SQLQ = SQLQ & nFld & "=LEFT(" & nFld & ", LEN(" & nFld & ")-3)+'" & clpNewCode & "' "
SQLQ = SQLQ & " WHERE "
SQLQ = SQLQ & " RIGHT(" & nFld & ", 3)= '" & clpOldCode & "'"

gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub UpdateTermEMPNBR(nTable As String, nFld As String, nFldTitle)
Dim SQLQ
Dim xTable
Select Case nTable
    Case "HR_ATTENDANCE_HISTORY", "HR_FOLLOW_UP", "HR_SECURE_BASIC", "HREMPEQU"
        Exit Sub
    Case "HR_OCC_HEALTH_SAFETY", "HR_COUNSEL"
        xTable = "Term_" & nTable
    Case "HRDOLENT", "HREARN", "HREDU", "HREMPSKL", "HRENTHRS", "HRTRADE"
        xTable = "Term_" & Mid(nTable, 3)
    Case Else
        If Mid(nTable, 3, 1) = "_" Then
          xTable = "Term" & Mid(nTable, 3)
        Else
          xTable = "Term_" & nTable
        End If
End Select

SQLQ = "UPDATE " & xTable & " SET "
SQLQ = SQLQ & nFld & "=LEFT(" & nFld & ", LEN(" & nFld & ")-3)+'" & clpNewCode & "' "
SQLQ = SQLQ & " WHERE "
SQLQ = SQLQ & " RIGHT(" & nFld & ", 3)= '" & clpOldCode & "'"

gdbAdoIhr001X.Execute SQLQ
End Sub

Private Function modAUDIT(nType As String, ACTX As String)
Dim TA As New ADODB.Recordset
Dim TB As New ADODB.Recordset
Dim xPT, xDiv, xADD
Dim TC As New ADODB.Recordset
Dim SQLQ
Dim nAuField, nAuOLDField, nTable
Dim nEmpField, nCHGField
'''On Error GoTo AUDIT_ERR
modAUDIT = False

nAuOLDField = ""

Select Case nType
Case "BNCD"
    nTable = "HRBENFT"
    nEmpField = "BF_EMPNBR"
    nAuField = "AU_BCODE"
    nCHGField = "BF_BCODE"
Case "EDOL"
    nTable = "HRDOLENT"
    nEmpField = "DE_EMPNBR"
    nAuField = "AU_DOLENT"
    nCHGField = "DE_TYPE"
Case "SDPP"
    nTable = "HR_SALARY_HISTORY"
    nEmpField = "SH_EMPNBR"
    nAuField = "AU_PAYP"
    nAuOLDField = "AU_OLDPAYP"
    nCHGField = "SH_PAYP"
Case "EARN"
    nTable = "HREARN"
    nEmpField = "EMPNBR"
    nAuField = "AU_EARN"
    nCHGField = "EARN_TYPE"
Case Else
    nTable = "HREMP"
    nEmpField = "ED_EMPNBR"
    Select Case nType
    Case "EMPNO"
        nAuField = "AU_EMPNBR"
        nCHGField = "ED_EMPNBR"
    Case "DEPT"
        nAuField = "AU_DEPTNO"
        nAuOLDField = "AU_OLDDEPT"
        nCHGField = "ED_DEPTNO"
    Case "DIV"
        nAuField = "AU_DIV"
        nAuOLDField = "AU_OLDDIV"
        nCHGField = "ED_DIV"
    Case "GLNO"
        nAuField = "AU_DEPT_GL"
        nCHGField = "ED_GLNO"
    Case "SALDIST"
        nAuField = "AU_SALDIST"
        nCHGField = "ED_SALDIST"
    Case "EDAB"
        nAuField = "AU_ADMINBY"
        nCHGField = "ED_ADMINBY"
    Case "EDEM"
        nAuField = "AU_EMP"
        nCHGField = "ED_EMP"
    Case "EDLC"
        nAuField = "AU_LOC"
        nAuOLDField = "AU_OLDLOC"
        nCHGField = "ED_LOC"
    Case "EDOR"
        nAuField = "AU_ORG"
        nCHGField = "ED_ORG"
    Case "EDRG"
        nAuField = "AU_REGION"
        nCHGField = "ED_REGION"
    Case "EDSE"
        nAuField = "AU_SECTION"
        nCHGField = "ED_SECTION"
    Case "EDSP"
        nAuField = "AU_SUPCODE"
        nCHGField = "ED_SUPCODE"
    Case "HMLN"
        nAuField = "AU_HOMELINE"
        nCHGField = "ED_HOMELINE"
    Case "HMOP"
        nAuField = "AU_HOMEOPRTNBR"
        nCHGField = "ED_HOMEOPRTNBR"
    Case "HMSF"
        nAuField = "AU_HOMESHIFT"
        nCHGField = "ED_HOMESHIFT"
    Case "HMWC"
        nAuField = "AU_HOMEWRKCNT"
        nCHGField = "ED_HOMEWRKCNT"
    Case "EDL1"
        nAuField = "AU_LANG1"
        'nCHGField = "ED_LANG" 'George Apr 4,2006 #10574
        nCHGField = "EL_LANG_SPOKEN" 'George Apr 4,2006 #10574
    Case Else
        modAUDIT = True
        Exit Function
    End Select
    
End Select
SQLQ = "INSERT INTO HRAUDIT ("
SQLQ = SQLQ & " AU_COMPNO"
SQLQ = SQLQ & ",AU_LDATE"
SQLQ = SQLQ & ",AU_LUSER"
SQLQ = SQLQ & ",AU_LTIME"
SQLQ = SQLQ & ",AU_UPLOAD"
SQLQ = SQLQ & ",AU_TYPE"
SQLQ = SQLQ & ",AU_NEWEMP"

SQLQ = SQLQ & ",AU_PTUPL"
SQLQ = SQLQ & ",AU_DIVUPL"

SQLQ = SQLQ & ",AU_EMPNBR"

If nType <> "EMPNO" Then SQLQ = SQLQ & "," & nAuField

If Len(nAuOLDField) > 0 Then SQLQ = SQLQ & "," & nAuOLDField

SQLQ = SQLQ & " )"

'Hemu - 06/12/2003 Begin - Since HRAUDIT is in different database
If (Not glbSQL) And (Not glbOracle) Then SQLQ = SQLQ & " IN '" & glbIHRAUDIT & "' [;PWD=petman;DATABASE=" & glbIHRAUDIT & "] "
'Hemu - 06/12/2003 End

SQLQ = SQLQ & " SELECT"
SQLQ = SQLQ & " '001'"
SQLQ = SQLQ & "," & Date_SQL(Date)
SQLQ = SQLQ & ",'" & glbUserID & "'"
SQLQ = SQLQ & ",'" & Time$ & "'"
SQLQ = SQLQ & ",'N'"
SQLQ = SQLQ & ",'" & ACTX & "'"
SQLQ = SQLQ & ",'N'"

SQLQ = SQLQ & IIf(nType = "EDPT", ",'" & clpNewCode & "'", ",ED_PT")
SQLQ = SQLQ & IIf(nType = "DIV", ",'" & clpNewCode & "'", ",ED_DIV")

If nType = "EMPNO" And glbLinamar Then
    SQLQ = SQLQ & ",LEFT(" & nCHGField & ",LEN(" & nCHGField & ")-3)+'" & clpNewCode & "'"
Else
    SQLQ = SQLQ & "," & nEmpField
End If


If nType <> "EMPNO" Then
    If optCodes(1) And glbLinamar Then
        SQLQ = SQLQ & ",'" & clpNewCode & "'+SUBSTRING(" & nCHGField & ",4,50)"
    Else
        If ACTX = "D" Then
            SQLQ = SQLQ & ",''"
        Else
            If glbLinamar And chkSep.Visible = True And chkSep <> 0 Then
                SQLQ = SQLQ & ",RIGHT(ED_EMPNBR,3)+'" & clpOldCode & "'"
            Else
                SQLQ = SQLQ & ",'" & clpNewCode & "'"
            End If
        End If
    End If
End If

If Len(nAuOLDField) > 0 Then SQLQ = SQLQ & ",'" & clpOldCode & "'"

SQLQ = SQLQ & " FROM HREMP"
If nTable <> "HREMP" Then
    If glbOracle Then
        SQLQ = SQLQ & "," & nTable & " WHERE  " & nTable & "." & nEmpField & "=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " AND "
    Else
        SQLQ = SQLQ & " INNER JOIN " & nTable & " ON  " & nTable & "." & nEmpField & "=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE "
    End If
Else
    SQLQ = SQLQ & " WHERE "
End If
SQLQ = SQLQ & glbSeleDeptUn
'Language isn't in HREMp anymore
If nType = "EDL1" And nCHGField = "EL_LANG_SPOKEN" Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN ( SELECT EL_EMPNBR FROM HR_LANGUAGE WHERE " & nCHGField & "='" & IIf(clpDIV.Visible, clpDIV, "") & clpOldCode & "')"
Else
    SQLQ = SQLQ & " AND " & nCHGField & "='" & IIf(clpDIV.Visible, clpDIV, "") & clpOldCode & "'"
End If

If glbWFC And Len(clpCode(0).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_SECTION = '" & clpCode(0).Text & "' "
End If
gdbAdoIhr001.Execute SQLQ
If glbWFC Then
    Call GetPayID(Date, Date)
End If

modAUDIT = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

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
UpdateRight = True
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
If optCodes(3) Then
    Deleteble = False
ElseIf optCodes(0) Or optCodes(1) Or optCodes(4) Or (optCodes(2) And glbLinamar) Then
    Deleteble = False
ElseIf optCodes(6) Or optCodes(7) Then
    Deleteble = False
Else
    Deleteble = True
End If
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Private Sub txtCourseName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
