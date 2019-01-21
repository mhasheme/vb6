VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVACSICK 
   Appearance      =   0  'Flat
   Caption         =   "Vacation and Sick Entitlements"
   ClientHeight    =   8040
   ClientLeft      =   75
   ClientTop       =   1380
   ClientWidth     =   11265
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
   ScaleHeight     =   8040
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fvasick.frx":0000
      Height          =   2535
      Left            =   90
      OleObjectBlob   =   "fvasick.frx":0014
      TabIndex        =   30
      Top             =   120
      Width           =   10605
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   7380
      Width           =   11265
      _Version        =   65536
      _ExtentX        =   19870
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
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "&Batch Recalculate"
         Height          =   375
         Index           =   1
         Left            =   9600
         TabIndex        =   13
         Tag             =   "Recalculate for all employees"
         Top             =   0
         Visible         =   0   'False
         Width           =   1785
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   6360
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
      Begin VB.CommandButton cmdFullTable 
         Appearance      =   0  'Flat
         Caption         =   "&Multiple Employee Edit"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Tag             =   "Multiple Employee Edit"
         Top             =   0
         Width           =   2475
      End
      Begin VB.CommandButton CmdRecalc 
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate 1 Employee"
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   12
         Tag             =   "Print the Employee Listing"
         Top             =   0
         Width           =   2295
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8310
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Tag             =   "Find Employee"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtEESearch 
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
      Left            =   2040
      TabIndex        =   8
      Tag             =   "00-Search for Surname"
      Top             =   4960
      Width           =   1935
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Employee Number"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   4920
      Width           =   2415
   End
   Begin Threed.SSPanel panDetails 
      Height          =   1425
      Left            =   0
      TabIndex        =   16
      Top             =   3240
      Width           =   11160
      _Version        =   65536
      _ExtentX        =   19685
      _ExtentY        =   2514
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
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ED_LUSER"
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
         Index           =   2
         Left            =   10140
         MaxLength       =   25
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "LUser"
         Top             =   960
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ED_LTIME"
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
         Index           =   1
         Left            =   10140
         MaxLength       =   25
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "LTime"
         Top             =   360
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         DataField       =   "ED_LDATE"
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
         Left            =   10140
         MaxLength       =   25
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Ldate"
         Top             =   600
         Visible         =   0   'False
         Width           =   650
      End
      Begin MSMask.MaskEdBox OutSick 
         Height          =   285
         Left            =   9120
         TabIndex        =   7
         ToolTipText     =   "Current Outstanding"
         Top             =   900
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox OutVac 
         Height          =   285
         Left            =   9120
         TabIndex        =   3
         ToolTipText     =   "Current Outstanding"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medSICKT 
         DataField       =   "ED_SICKT"
         Height          =   285
         Left            =   6360
         TabIndex        =   6
         Tag             =   "11-Total number of hours Sick time booked and taken"
         ToolTipText     =   "Total Taken and Booked"
         Top             =   900
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medVacT 
         DataField       =   "ED_VACT"
         Height          =   285
         Left            =   6360
         TabIndex        =   2
         Tag             =   "11-Total number of hours Vacation time booked and taken"
         ToolTipText     =   "Total Taken and Booked"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPSick 
         DataField       =   "ED_PSICK"
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Tag             =   "11-Banked hours of Sick time from previous year"
         Top             =   900
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCSick 
         DataField       =   "ED_SICK"
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Tag             =   "11-Total Number of hours of Sick time entitled"
         Top             =   900
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medPVac 
         DataField       =   "ED_PVAC"
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Tag             =   "11-Banked hours of Vacation from previous year"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox medCVac 
         DataField       =   "ED_VAC"
         Height          =   285
         Left            =   3840
         TabIndex        =   1
         Tag             =   "11-Total number of hours Vacation time entitled"
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "Fixed"
         PromptChar      =   "_"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "VACATION"
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
         Left            =   120
         TabIndex        =   29
         Top             =   525
         Width           =   975
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   28
         Top             =   525
         Width           =   840
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SICKTIME"
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
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   945
         Width           =   855
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Previous"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   26
         Top             =   945
         Width           =   810
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taken && Booked"
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
         Left            =   5040
         TabIndex        =   25
         Top             =   525
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Taken && Booked"
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
         Left            =   5040
         TabIndex        =   24
         Top             =   945
         Width           =   1200
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Outstanding"
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
         Left            =   7560
         TabIndex        =   23
         Top             =   525
         Width           =   1410
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Outstanding"
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
         Left            =   7560
         TabIndex        =   22
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   3000
         TabIndex        =   21
         Top             =   525
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Current"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   3000
         TabIndex        =   20
         Top             =   945
         Width           =   735
      End
   End
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Surname      "
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
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   1425
   End
End
Attribute VB_Name = "frmVACSICK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EESNameSort As Integer
Dim oPVac, oPSick, oVac, oSick
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew


Private Sub CalcOutS()
OutVac = 0

If IsNumeric(Data1.Recordset("ED_VAC")) Then
   OutVac = OutVac + Data1.Recordset("ED_VAC")
End If

If IsNumeric(Data1.Recordset("ED_PVAC")) Then
   OutVac = OutVac + Data1.Recordset("ED_PVAC")
End If

If IsNumeric(Data1.Recordset("ED_VACT")) Then
   OutVac = OutVac - Data1.Recordset("ED_VACT")
End If

OutSick = 0

If IsNumeric(Data1.Recordset("ED_SICK")) Then
   OutSick = OutSick + Data1.Recordset("ED_SICK")
End If

If IsNumeric(Data1.Recordset("ED_PSICK")) Then
   OutSick = OutSick + Data1.Recordset("ED_PSICK")
End If

If IsNumeric(Data1.Recordset("ED_SICKT")) Then
   OutSick = OutSick - Data1.Recordset("ED_SICKT")
End If

End Sub

Private Function chkVac()
Dim dd As Integer

chkVac = False
If Len(medCVac) < 1 Then
    MsgBox "Current Vacation must be entered"
    medCVac.SetFocus
    Exit Function
End If

If Not IsNumeric(medCVac) Then
    MsgBox "Invalid Current Vacation"
    medCVac.SetFocus
    Exit Function
End If

If Len(medPVac) < 1 Then
    medPVac = 0
End If

If Not IsNumeric(medPVac) Then
    MsgBox "Invalid Previous Vacation"
    medCVac.SetFocus
    Exit Function
End If

If Len(medCSick) < 1 Then
    MsgBox "Current Sick Time must be entered"
    medCSick.SetFocus
    Exit Function
End If

If Not IsNumeric(medCSick) Then
    MsgBox "Invalid Current Sick Time"
    medCSick.SetFocus
    Exit Function
End If

If Len(medPSick) < 1 Then
    medPSick = 0
End If

If Not IsNumeric(medPSick) Then
    MsgBox "Invalid Previous Sick Time"
    medPSick.SetFocus
    Exit Function
End If

chkVac = True

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
''' Sam add July 2002 * Remove Binding Control
If rsDATA.State <> 0 Then
    rsDATA.CancelUpdate
End If
Call Display_Value



Call modSTUPD(True)  ' reset screen's attributes


vbxTrueGrid.Enabled = True
vbxTrueGrid.MarqueeStyle = 4
vbxTrueGrid.EditActive = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '28July99 js

End Sub

Sub cmdClose_Click()
    Call NextForm
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdEESort_Click()

txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort.Caption = "Sort by Surname "
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort.Caption = "Sort by Emp. #"
End If

If EERetrieve() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = " "
txtEESearch.SetFocus

End Sub

Private Sub cmdEESort_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark

On Error GoTo Srch_Err

If Not Len(txtEESearch) > 0 Then
   MsgBox "To search you must enter something to search for."
   Exit Sub
End If
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch, "'", "''")
    If EESNameSort = True Then
        SQLQ = "ED_SURNAME  >= '" & Sch & "'"
    Else
        If Not IsNumeric(txtEESearch.Text) And Not glbLinamar Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        If glbLinamar Then
            SQLQ = "EMPNBR >= '" & Sch & "'"
        Else
            SQLQ = "ED_EMPNBR >= '" & Sch & "'"
        End If
        Data1.Recordset.Find SQLQ
    End If
    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    MsgBox "Employee not found"
    Data1.Refresh
End If



Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "Find Next")
Call RollBack '28July99 js

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFullTable_Click()
Dim Msg$, DgDef As Variant, Response As Variant

If Not gSec_Upd_Entitlements Then
    Msg$ = "You do not have authority for this transaction"
    Exit Sub
End If

Call modSTUPD(True)
medCVac.SetFocus
'cmdCancel.Enabled = False
vbxTrueGrid.AllowUpdate = True
vbxTrueGrid.Enabled = True
vbxTrueGrid.EditActive = True
vbxTrueGrid.Refresh
vbxTrueGrid.MarqueeStyle = 2

End Sub

Private Sub cmdFullTable_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call modSTUPD(True)


medCVac.Enabled = True
medPVac.SetFocus


Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '28July99 js

End Sub

Sub cmdOK_Click()
Dim DtTm As Variant, rc As Integer
Dim x, xID

DtTm = Now

On Error GoTo Add_Err

If Not chkVac() Then Exit Sub
Call Trans_Accrual

glbENTScreen = True

Call UpdUStats(Me) ' update user's stats (who did it and when)

xID = Data1.Recordset("ED_EMPNBR")
rsDATA!ED_PVAC = Val(medPVac)
Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh


Data1.Recordset.Find "ED_EMPNBR=" & xID
Call SET_UP_MODE

vbxTrueGrid.EditActive = False
vbxTrueGrid.MarqueeStyle = 4
vbxTrueGrid.Enabled = True
vbxTrueGrid.AllowUpdate = False
Call CalcOutS

Dim xKey

xKey = Data1.Recordset("ED_EMPNBR")
xKey = xKey & "|" & Format(Data1.Recordset("ED_EFDATE"), "dd-mmm-yyyy")
xKey = xKey & "|" & Format(Data1.Recordset("ED_ETDATE"), "dd-mmm-yyyy")
xKey = xKey & "|VAC"
xKey = xKey & "|" & medCVac
xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
Call Entitlements_Master_Integration(xKey, xID) 'George added for Advance Tracker

xKey = Data1.Recordset("ED_EMPNBR")
xKey = xKey & "|" & Format(Data1.Recordset("ED_EFDATES"), "dd-mmm-yyyy")
xKey = xKey & "|" & Format(Data1.Recordset("ED_ETDATES"), "dd-mmm-yyyy")
xKey = xKey & "|SICK"
xKey = xKey & "|" & medCSick
xKey = xKey & "|" & Format(Date, "dd-mmm-yyyy") 'Transaction Date
Call Entitlements_Master_Integration(xKey, xID) 'George added for Advance Tracker

Call NextForm
Exit Sub

Add_Err:
If Err = 3197 Then Resume Next
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack '28July99 js

Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport

RHeading = "Vacation and Sick Entitlements Listing"

'----------\\
    RHeading = Me.Caption
    Me.vbxCrystal.Reset
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading

    xReport = glbIHRREPORTS & "rgvacsi1.rpt"

    
    Me.vbxCrystal.ReportFileName = xReport
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    If EESNameSort = True Then  ' was sorted by surname
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_SURNAME}"
        Me.vbxCrystal.SortFields(1) = "+{HREMP.ED_FNAME}"
    Else
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_EMPNBR}"
    End If
    
    ' dkostka - 10/18/2001 - Added check for security, used to print for all facilities.
    glbiOneWhere = False
    glbstrSelCri = ""
    glbCri_DeptUN ""
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport

RHeading = "Vacation and Sick Entitlements Listing"

'----------\\
    RHeading = Me.Caption
    Me.vbxCrystal.Reset
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading

    xReport = glbIHRREPORTS & "rgvacsi1.rpt"

    
    Me.vbxCrystal.ReportFileName = xReport
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDB
    'End If
    If EESNameSort = True Then  ' was sorted by surname
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_SURNAME}"
        Me.vbxCrystal.SortFields(1) = "+{HREMP.ED_FNAME}"
    Else
        Me.vbxCrystal.SortFields(0) = "+{HREMP.ED_EMPNBR}"
    End If
    
    ' dkostka - 10/18/2001 - Added check for security, used to print for all facilities.
    glbiOneWhere = False
    glbstrSelCri = ""
    glbCri_DeptUN ""
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdRecalc_Click(Index As Integer)
Dim Msg, Response, DgDef, SQLQ As String

Msg = "Do you wish to proceed and recalculate "
If Index = 1 Then
    Msg = Msg & "all Employees' "
Else
    Msg = Msg & "the Employee's "
End If
Msg = Msg & "outstanding entitlement ?"

'Ticket #20020
If glbEntOutStanding$ <> "1" Or glbEntOutStandingS$ <> "1" Then
    Msg = Msg & vbCrLf & vbCrLf & "NOTE: If the Entitlement Date Range of an employee has ended prior to today's date, "
    Msg = Msg & vbCrLf & "then the Entitlement Period for that employee will be change to new Entitlement Date Range."
End If

DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, "ReCalculate")
If Response = IDNO Then Exit Sub

Screen.MousePointer = HOURGLASS
If Index = 1 Then
    Call EntReCalc("")
Else
    If glbGuelph Then   ' FOR Guelph-Willington
        Call AddFTE(Data1.Recordset("ED_EMPNBR"), "NEW")
    End If
    SQLQ = "ED_EMPNBR = " & Data1.Recordset("ED_EMPNBR")
    
    'County of Essex - Ticket #12676
    If glbCompSerial = "S/N - 2192W" Then
        Call EntReCalc(SQLQ, True)
    Else
        Call EntReCalc(SQLQ)
    End If
    
    If glbCompSerial = "S/N - 2173W" Then 'Town of Ajax 'Ticket #30402 Franks 08/02/2017
        Call Recalculate_OTBANK_Ajax_AllEmployees(Data1.Recordset("ED_EMPNBR"))
    End If
End If
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Screen.MousePointer = DEFAULT

glbENTScreen = True
If Index = 1 Then
    Call Form_Activate
Else
    Data1.Recordset.Find SQLQ
End If

'Call VacSickHourlyFollowUp("VAC", Date)
'Call VacSickHourlyFollowUp("SICK", Date)
glbflgFU = False

End Sub
Function EERetrieve()

Dim SQLQ As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

On Error GoTo EERetrieve_Err

EERetrieve = False         ' if not found - no depts
Screen.MousePointer = HOURGLASS

SQLQ = "Select ED_SURNAME, ED_FNAME ,"
If glbLinamar Then
    SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
Else
    If glbOracle Then
        SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
    Else
        SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
    End If
    
End If
SQLQ = SQLQ & "ED_EMPNBR, ED_VAC, ED_PVAC,ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES, "
SQLQ = SQLQ & "ED_SICK, ED_PSICK, ED_LDATE, ED_LTIME, ED_LUSER, "
SQLQ = SQLQ & "ED_VACT, ED_SICKT "
If glbtermopen Then
    SQLQ = SQLQ & ", TERM_SEQ "
    SQLQ = SQLQ & " FROM Term_HREMP "
Else
    SQLQ = SQLQ & " FROM HREMP "
End If

SQLQ = SQLQ & " Where " & glbSeleDeptUn

If EESNameSort = True Then
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
Else
    SQLQ = SQLQ & " ORDER BY " & IIf(glbLinamar, "EMPNBR", "ED_EMPNBR")
End If
    
Data1.RecordSource = SQLQ
Data1.Refresh
Me.vbxTrueGrid.Refresh
If glbtermopen Then
    If glbTERM_Seq > 0 Then
        SQLQ = "TERM_SEQ = " & glbTERM_Seq
        Data1.Recordset.Find SQLQ
    End If
Else
    If glbLEE_ID > 0 Then
        SQLQ = "ED_EMPNBR = " & glbLEE_ID
        Data1.Recordset.Find SQLQ
    End If
End If
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieve_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "Select")
Call RollBack '28July99 js

End Function

Private Sub CmdRecalc_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Dim SQLQ
glbOnTop = "FRMVACSICK"
If glbENTScreen = True Then
    glbENTScreen = False
    If EERetrieve() = False Then     ' get the info for this person
        Exit Sub
    End If
    If glbtermopen Then
        If glbTERM_Seq > 0 Then
           SQLQ = "TERM_SEQ = " & glbTERM_Seq
           Data1.Recordset.Find SQLQ
        End If
    Else
        If glbLEE_ID > 0 Then
           SQLQ = "ED_EMPNBR = " & glbLEE_ID
           Data1.Recordset.Find SQLQ
        End If
    End If
End If
fglbNew = False
Call SET_UP_MODE

End Sub

Private Sub Form_Deactivate()
    MDIMain.panHelp(0).Caption = "info:HR Main functions Locked until EE Selected"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMVACSICK"
End Sub

Private Sub Form_Load()
Dim SQLQ As String

glbOnTop = "FRMVACSICK"
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).Caption = "Retrieving Employee List - Stand by"
EESNameSort = True  'first sort is by surname

Screen.MousePointer = DEFAULT

glbENTScreen = True

If glbCompSerial = "S/N - 2262W" Then
    medPVac.Format = "#,##0.0000"
    medCVac.Format = "#,##0.0000"
    medVacT.Format = "#,##0.0000"
    OutVac.Format = "#,##0.0000"
    vbxTrueGrid.Columns(3).NumberFormat = "0.0000"
    vbxTrueGrid.Columns(4).NumberFormat = "0.0000"
    vbxTrueGrid.Columns(5).NumberFormat = "0.0000"
End If

Call modSTUPD(True)
If Not gSec_Upd_Entitlements Then
'    cmdModify.Enabled = False
    cmdFullTable.Enabled = False
    CmdRecalc(0).Enabled = False
'    cmdRecalc(1).Enabled = False
End If
Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

If glbWFC Then
    lblTitle(2).Visible = False
    lblTitle(3).Visible = False
    lblTitle(5).Visible = False
    lblTitle(7).Visible = False
    lblTitle(9).Visible = False
    medPSick.Visible = False
    medCSick.Visible = False
    medSICKT.Visible = False
    OutSick.Visible = False
    Me.Caption = "Vacation Entitlements"
    vbxTrueGrid.Columns(6).Visible = False
    vbxTrueGrid.Columns(7).Visible = False
    vbxTrueGrid.Columns(8).Visible = False
End If

If glbCompSerial = "S/N - 2418W" Then
    lblTitle(2).Visible = False
    lblTitle(3).Visible = False
    lblTitle(5).Visible = False
    lblTitle(7).Visible = False
    lblTitle(9).Visible = False
    medPSick.Visible = False
    medCSick.Visible = False
    medSICKT.Visible = False
    OutSick.Visible = False
    Me.Caption = "Vacation Entitlements"
    vbxTrueGrid.Columns(6).Visible = False
    vbxTrueGrid.Columns(7).Visible = False
    vbxTrueGrid.Columns(8).Visible = False
End If

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
    MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
    Call NextForm
End Sub


Private Sub medCSick_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medCVac_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub medPSick_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub medPVac_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

panDetails.Enabled = TF
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdModify.Enabled = FT
cmdFullTable.Enabled = TF 'FT
CmdRecalc(0).Enabled = TF  'FT
CmdRecalc(1).Enabled = TF  'FT
'cmdClose.Enabled = FT
'cmdPrint.Enabled = FT
cmdEESort.Enabled = TF    'FT
txtEESearch.Enabled = TF   'FT
cmdFind.Enabled = TF     'FT
medCSick.Enabled = TF
medCVac.Enabled = TF
medPSick.Enabled = TF
medPVac.Enabled = TF
'medSICKT.Enabled = TF
'medVacT.Enabled = TF
' vbxTrueGrid.Enabled = FT
If glbtermopen Then
    CmdRecalc(0).Visible = False
    CmdRecalc(1).Visible = False
End If
End Sub

Private Sub medSICKT_LostFocus()
If Not IsNumeric(medSICKT) Then medSICKT = Val(medSICKT)
End Sub

Private Sub medVacT_LostFocus()
    If Not IsNumeric(medVacT) Then medVacT = Val(medVacT)
End Sub



Private Sub txtEESearch_GotFocus()
    Call SetPanHelp(ActiveControl)
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
        
        SQLQ = "Select ED_SURNAME, ED_FNAME ,"
        If glbLinamar Then
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        Else
            If glbOracle Then
                SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
            Else
                SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
            End If
            
        End If
        SQLQ = SQLQ & "ED_EMPNBR, ED_VAC, ED_PVAC,ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES, "
        SQLQ = SQLQ & "ED_SICK, ED_PSICK, ED_LDATE, ED_LTIME, ED_LUSER, "
        SQLQ = SQLQ & "ED_VACT, ED_SICKT "
        If glbtermopen Then
            SQLQ = SQLQ & ", TERM_SEQ "
            SQLQ = SQLQ & " FROM Term_HREMP "
        Else
            SQLQ = SQLQ & " FROM HREMP "
        End If
        
        SQLQ = SQLQ & " Where " & glbSeleDeptUn
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    cmdOK.SetFocus
End If

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




Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Data1.Recordset.EOF Then
    Call CalcOutS
    Call Display_Value
End If
End Sub


Sub AddFTE(xEmpNo, xFLAG)
    Dim OldFTE, NewFTE, xEFDATE, xETDATE, xNumVac
    Dim fNewFTE, fOldFTE, FlagOldFTE
    Dim RsFTEHis As New ADODB.Recordset
    Dim xDays1, xDays2, xVacDays, xDate1, xDate2, xFDate, xTDate, xHrsDay, xHrsDayN
    Dim xVacHours, xYear, xNum As Integer, II, J
    Dim xArray(100, 2)
    Dim tNewFTE, xNumVacINS, VAC_First
    Dim RsTempEmp As New ADODB.Recordset
    Dim RsJobEmp As New ADODB.Recordset
    Dim SQLQ, xTxtJOB
    Dim FlagLoop As Boolean
    
    SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
    RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xEFDATE = ""
    xETDATE = ""
    xNumVac = 0
    If Not RsTempEmp.EOF Then
        xNumVac = RsTempEmp("ED_VAC")
        xNumVacINS = RsTempEmp("ED_VAC")
        xEFDATE = RsTempEmp("ED_EFDATE")
        xETDATE = RsTempEmp("ED_ETDATE")
    End If
    RsTempEmp.Close
    
    If Len(xEFDATE) = 0 Or Len(xETDATE) = 0 Then
        Exit Sub
    End If
    
    SQLQ = "Select * from HR_JOB_HISTORY Where JH_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
    RsJobEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If RsJobEmp.EOF Then
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM FTE_HISTORY WHERE CP_EMPNBR = " & xEmpNo & " "
    If IsDate(xEFDATE) Then
        SQLQ = SQLQ & "AND CP_FDATE = " & Date_SQL(xEFDATE)
    End If
    If IsDate(xETDATE) Then
        SQLQ = SQLQ & "AND CP_TDATE = " & Date_SQL(xETDATE)
    End If
    SQLQ = SQLQ & "ORDER BY CP_FDATE DESC"
    RsFTEHis.Open SQLQ, gdbAdoSN2322, adOpenKeyset, adLockOptimistic
    If RsFTEHis.EOF And xFLAG <> "NEW" Then
        Exit Sub
    End If

    If xFLAG = "NEW" Then
        If xNumVac = 0 Then
            Exit Sub
        End If
        If Not RsFTEHis.EOF Then ' IF CP_VACORIGION EXIST AND CHANGE IN THE SAME YEAR
            If RsFTEHis("CP_FDATE") = xEFDATE Then
                xNumVac = RsFTEHis("CP_VACORIGION")
                GoTo MAIN_DEAL
            End If
        End If
        '' The following shows how to calculate the VAC days at the end of last year
        '' We always suppose the FTE# is 1.00 at the end of last year
        ' X is VAC days when FTE# = 1
        ' VAC_First is the first VAC days before FTE# change
        ' days1,days2, ... daysn are date range when FTE# change within this year
        ' VAC_First = X/365 * FTE#1 * days1 + X/365 * FTE#2 * days2 + ... + X/365 * FTE#n * daysn
        ' X = (VAC_First * 365)/(FTE#1 * days1 + FTE#2 * days2 + ... + FTE#n * daysn)
        VAC_First = xNumVac
        
        xDate1 = "**"
        xFDate = xEFDATE
        xTDate = xETDATE
        FlagLoop = True
        xHrsDayN = 0
        If RsJobEmp("JH_DHRS") = 0 Then
            xHrsDayN = 0
        Else
            If IsNull(RsJobEmp("JH_DHRS")) Then
                xHrsDayN = 0
            Else
                xHrsDayN = RsJobEmp("JH_DHRS")
            End If
        End If
        If IsNull(RsJobEmp("JH_FTENUM")) Then
            fNewFTE = 0
        Else
            fNewFTE = RsJobEmp("JH_FTENUM")
        End If
        RsJobEmp.MoveNext
        fOldFTE = 0
        FlagOldFTE = True
        II = 0
        Do While (Not RsJobEmp.EOF) And FlagLoop
            xDate1 = RsJobEmp("JH_SDATE")
            If FlagOldFTE Then
                If Not IsNull(RsJobEmp("JH_FTENUM")) Then
                    fOldFTE = RsJobEmp("JH_FTENUM")
                End If
                FlagOldFTE = False
            End If
            If CVDate(xDate1) > CVDate(xETDATE) Then
                GoTo Next_Rec00
            End If
            If RsJobEmp("JH_FTENUM") = 0 Then
                GoTo Next_Rec00
            End If
            If IsNull(RsJobEmp("JH_FTENUM")) Then
                GoTo Next_Rec00
            End If
            OldFTE = RsJobEmp("JH_FTENUM")
            
            If RsJobEmp("JH_DHRS") = 0 Then
                GoTo Next_Rec00
            End If
            If IsNull(RsJobEmp("JH_DHRS")) Then
                GoTo Next_Rec00
            End If
            xHrsDay = RsJobEmp("JH_DHRS")
            
            If CVDate(xDate1) < CVDate(xEFDATE) Then
                II = II + 1
                xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate)) * OldFTE
                FlagLoop = False
            Else
                II = II + 1
                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
                xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
            End If
            
Next_Rec00:
            RsJobEmp.MoveNext
        Loop
        If IsDate(xDate1) Then
            If CVDate(xDate1) > CVDate(xEFDATE) Then
                II = II + 1
                xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate)) * OldFTE
            End If
        End If
        
        xVacDays = 0
        For J = 1 To II
            xVacDays = xVacDays + xArray(J, 1)
        Next
        If xVacDays = 0 Then
            Exit Sub
        End If
        If xHrsDay = 0 Then
            Exit Sub
        End If
        xNumVac = Round((((VAC_First * 365) / (xVacDays)) / xHrsDayN), 0) * xHrsDayN

    End If
        
   
    '--- Above Got vacation days per year when FTE = 1 (xNumVac)
MAIN_DEAL:
    II = 0
    xDate1 = "**"
    xFDate = xEFDATE
    xTDate = xETDATE
    FlagLoop = True
    RsJobEmp.MoveFirst
    Do While (Not RsJobEmp.EOF) And FlagLoop
        xDate1 = RsJobEmp("JH_SDATE")
        If CVDate(xDate1) > CVDate(xETDATE) Then
            GoTo Next_Rec01
        End If
        If RsJobEmp("JH_FTENUM") = 0 Then
            GoTo Next_Rec01
        End If
        If IsNull(RsJobEmp("JH_FTENUM")) Then
            GoTo Next_Rec01
        End If
        OldFTE = RsJobEmp("JH_FTENUM")
        
        If RsJobEmp("JH_DHRS") = 0 Then
            GoTo Next_Rec01
        End If
        If IsNull(RsJobEmp("JH_DHRS")) Then
            GoTo Next_Rec01
        End If
        xHrsDay = RsJobEmp("JH_DHRS")
        
        If CVDate(xDate1) < CVDate(xEFDATE) Then
            II = II + 1
            xArray(II, 1) = DateDiff("d", CVDate(xFDate), CVDate(xTDate))
            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
            FlagLoop = False
        Else
            II = II + 1
            xArray(II, 1) = DateDiff("d", CVDate(xDate1), CVDate(xTDate))
            xArray(II, 2) = xArray(II, 1) * Round(((xNumVac * OldFTE) / (365 * xHrsDay)), 3)
            xTDate = xDate1 'DateAdd("d", -1, CVDate(xDate1))
            
        End If
        
Next_Rec01:
        RsJobEmp.MoveNext
    Loop
    
    xVacDays = 0
    For J = 1 To II
        xVacDays = xVacDays + xArray(J, 2)
    Next
    
    If xVacDays = 0 Then
        Exit Sub
    End If
    xVacHours = Round(xVacDays, 0) * xHrsDay
    
    If xVacHours <> xNumVacINS Then
        gdbAdoIhr001.BeginTrans
        'Dim RsTempEmp As New ADODB.Recordset
        SQLQ = "Select ED_EMPNBR,ED_VAC,ED_EFDATE,ED_ETDATE from HREMP Where ED_EMPNBR = " & xEmpNo
        RsTempEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        If Not RsTempEmp.EOF Then
            RsTempEmp("ED_VAC") = xVacHours
            RsTempEmp.Update
        End If
        RsTempEmp.Close
        gdbAdoIhr001.CommitTrans
        
        If RsFTEHis.EOF Then
            RsFTEHis.AddNew
            RsFTEHis("CP_EMPNBR") = xEmpNo
            RsFTEHis("CP_VACORIGION") = xNumVac
            RsFTEHis("CP_VACO") = xNumVacINS
            RsFTEHis("CP_VACN") = xVacHours
            If fOldFTE > 0 Then
            RsFTEHis("CP_FTENUMO") = fOldFTE
            End If
            If fNewFTE > 0 Then
            RsFTEHis("CP_FTENUMN") = fNewFTE
            End If
            RsFTEHis("CP_FDATE") = CVDate(xEFDATE)
            RsFTEHis("CP_TDATE") = CVDate(xETDATE)
            RsFTEHis("CP_LDATE") = Date
            RsFTEHis("CP_LTIME") = Time$
            RsFTEHis("CP_LUSER") = glbUserID
            RsFTEHis.Update
        End If
    End If
    RsFTEHis.Close
    
    Exit Sub

ExitLin1:
End Sub

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
        'Me.cmdModify_Click
        Exit Sub
    End If
    
    
    SQLQ = "Select ED_SURNAME, ED_FNAME ,"
If glbLinamar Then
    SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
Else
    If glbOracle Then
        SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
    Else
        SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
    End If
End If
SQLQ = SQLQ & "ED_EMPNBR, ED_VAC, ED_PVAC, "
SQLQ = SQLQ & "ED_SICK, ED_PSICK, ED_LDATE, ED_LTIME, ED_LUSER, "
SQLQ = SQLQ & "ED_VACT, ED_SICKT "
If glbtermopen Then
    SQLQ = SQLQ & ", TERM_SEQ "
    SQLQ = SQLQ & " FROM Term_HREMP "
Else
    SQLQ = SQLQ & " FROM HREMP "
End If

SQLQ = SQLQ & " Where ED_EMPNBR= " & Data1.Recordset!ED_EMPNBR

If glbtermopen Then

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
 
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    oVac = medCVac
    oPVac = medPVac
    oSick = medCSick
    oPSick = medPSick

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
UpdateRight = gSec_Upd_Entitlements
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
    Updateble = True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
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
Call set_Buttons(UpdateState)
'Call modSTUPD(TF)
If Not UpdateRight Then TF = False
Call modSTUPD(TF)
End Sub

Private Sub Trans_Accrual()
Dim xDiffHours
Dim xEmpnbr
Dim xComments

If Data1.Recordset.EOF Then Exit Sub

xEmpnbr = Data1.Recordset!ED_EMPNBR

If Not IsNumeric(xEmpnbr) Then Exit Sub

'Ticket #23141 - They want to now pass today's date instead of Entitlement Start Date.
'If glbCompSerial = "S/N - 2363W" Then  'City of Kawartha Lakes
'    'They want to pass the Entitlement's Start Date instead of today's date
'    xDiffHours = Val(medPVac) - Val(oPVac)
'    'Ticket #22730
'    'xComments = "Prev. Vac. Ent. Chg from " & oPVac & " to " & medPVac
'    xComments = "Prev. Vac. Ent. Chg from " & oPVac & " to " & medPVac & ". OS: " & (Val(IIf(IsNull(oPVac), 0, oPVac)) + Val(IIf(IsNull(oVac), 0, oVac))) - Val(IIf(IsNull(medVacT), 0, medVacT))
'
'    Call Append_Accrual(xEmpnbr, "VAC", Data1.Recordset!ED_EFDATE, xDiffHours, "C", xComments)
'
'    xDiffHours = Val(medCVac) - Val(oVac)
'    'Ticket #22730
'    'xComments = "Current Vac. Ent. Chg from " & oVac & " to " & medCVac
'    xComments = "Current Vac. Ent. Chg from " & oVac & " to " & medCVac & ". OS: " & (Val(IIf(IsNull(oPVac), 0, oPVac)) + Val(IIf(IsNull(oVac), 0, oVac))) - Val(IIf(IsNull(medVacT), 0, medVacT))
'    Call Append_Accrual(xEmpnbr, "VAC", Data1.Recordset!ED_EFDATE, xDiffHours, "C", xComments)
'
'    xDiffHours = Val(medPSick) - Val(oPSick)
'    'Ticket #22730
'    'xComments = "Prev. Sick Ent. Chg from " & oPSick & " to " & medPSick
'    xComments = "Prev. Sick Ent. Chg from " & oPSick & " to " & medPSick & ". OS: " & (Val(IIf(IsNull(oPSick), 0, oPSick)) + Val(IIf(IsNull(oSick), 0, oSick))) - Val(IIf(IsNull(medSICKT), 0, medSICKT))
'    Call Append_Accrual(xEmpnbr, "SICK", Data1.Recordset!ED_EFDATES, xDiffHours, "C", xComments)
'
'    xDiffHours = Val(medCSick) - Val(oSick)
'    'Ticket #22730
'    'xComments = "Current Sick Ent. Chg from " & oSick & " to " & medCSick
'    xComments = "Current Sick Ent. Chg from " & oSick & " to " & medCSick & ". OS: " & (Val(IIf(IsNull(oPSick), 0, oPSick)) + Val(IIf(IsNull(oSick), 0, oSick))) - Val(IIf(IsNull(medSICKT), 0, medSICKT))
'    Call Append_Accrual(xEmpnbr, "SICK", Data1.Recordset!ED_EFDATES, xDiffHours, "C", xComments)
'Else
    'VACATION
    'Ticket #23357 Franks 02/28/2013
    If Not IsNull(Data1.Recordset!ED_EFDATE) And Not IsNull(Data1.Recordset!ED_ETDATE) Then
        'VACATION - Previous
        xDiffHours = Val(medPVac) - Val(oPVac)
        If xDiffHours <> 0 Then
            'Ticket #22730
            'xComments = "Prev. Vac. Ent. Chg from " & oPVac & " to " & medPVac
            xComments = "Prev. Vac. Ent. Chg from " & oPVac & " to " & medPVac & ". OS: " & (Val(IIf(IsNull(oPVac), 0, oPVac)) + Val(IIf(IsNull(oVac), 0, oVac))) - Val(IIf(IsNull(medVacT), 0, medVacT))
            If CVDate(Date) >= CVDate(Data1.Recordset!ED_EFDATE) And CVDate(Date) <= CVDate(Data1.Recordset!ED_ETDATE) Then
                Call Append_Accrual(xEmpnbr, "VAC", Date, xDiffHours, "C", xComments)
            Else
                Call Append_Accrual(xEmpnbr, "VAC", Data1.Recordset!ED_ETDATE, xDiffHours, "C", xComments)
            End If
        End If
        
        'VACATION - Current
        xDiffHours = Val(medCVac) - Val(oVac)
        If xDiffHours <> 0 Then
            'Ticket #22730
            'xComments = "Current Vac. Ent. Chg from " & oVac & " to " & medCVac
            xComments = "Current Vac. Ent. Chg from " & oVac & " to " & medCVac & ". OS: " & (Val(IIf(IsNull(oPVac), 0, oPVac)) + Val(IIf(IsNull(oVac), 0, oVac))) - Val(IIf(IsNull(medVacT), 0, medVacT))
            If CVDate(Date) >= CVDate(Data1.Recordset!ED_EFDATE) And CVDate(Date) <= CVDate(Data1.Recordset!ED_ETDATE) Then
                Call Append_Accrual(xEmpnbr, "VAC", Date, xDiffHours, "C", xComments)
            Else
                Call Append_Accrual(xEmpnbr, "VAC", Data1.Recordset!ED_ETDATE, xDiffHours, "C", xComments)
            End If
        End If
    End If
    
    'SICK - Previous
    'Ticket #23357 Franks 02/28/2013
    If Not IsNull(Data1.Recordset!ED_EFDATES) And Not IsNull(Data1.Recordset!ED_ETDATES) Then
        'SICK - Previous
        xDiffHours = Val(medPSick) - Val(oPSick)
        If xDiffHours <> 0 Then
            'Ticket #22730
            'xComments = "Prev. Sick Ent. Chg from " & oPSick & " to " & medPSick
            xComments = "Prev. Sick Ent. Chg from " & oPSick & " to " & medPSick & ". OS: " & (Val(IIf(IsNull(oPSick), 0, oPSick)) + Val(IIf(IsNull(oSick), 0, oSick))) - Val(IIf(IsNull(medSICKT), 0, medSICKT))
            If CVDate(Date) >= CVDate(Data1.Recordset!ED_EFDATES) And CVDate(Date) <= CVDate(Data1.Recordset!ED_ETDATES) Then
                Call Append_Accrual(xEmpnbr, "SICK", Date, xDiffHours, "C", xComments)
            Else
                Call Append_Accrual(xEmpnbr, "SICK", Data1.Recordset!ED_ETDATES, xDiffHours, "C", xComments)
            End If
        End If
        
        'SICK - Current
        xDiffHours = Val(medCSick) - Val(oSick)
        If xDiffHours <> 0 Then
            'Ticket #22730
            'xComments = "Current Sick Ent. Chg from " & oSick & " to " & medCSick
            xComments = "Current Sick Ent. Chg from " & oSick & " to " & medCSick & ". OS: " & (Val(IIf(IsNull(oPSick), 0, oPSick)) + Val(IIf(IsNull(oSick), 0, oSick))) - Val(IIf(IsNull(medSICKT), 0, medSICKT))
            If CVDate(Date) >= CVDate(Data1.Recordset!ED_EFDATES) And CVDate(Date) <= CVDate(Data1.Recordset!ED_ETDATES) Then
                Call Append_Accrual(xEmpnbr, "SICK", Date, xDiffHours, "C", xComments)
            Else
                Call Append_Accrual(xEmpnbr, "SICK", Data1.Recordset!ED_ETDATES, xDiffHours, "C", xComments)
            End If
        End If
    End If
'End If

End Sub

