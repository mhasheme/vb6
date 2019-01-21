VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmREEO 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "EEO Reports"
   ClientHeight    =   7440
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   10980
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
   ScaleHeight     =   7440
   ScaleWidth      =   10980
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTitle 
      Height          =   345
      Left            =   4080
      TabIndex        =   9
      Text            =   "txtTitle"
      Top             =   720
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame fraWorkForce 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   4215
      Begin VB.ComboBox comGroup 
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
         Index           =   1
         ItemData        =   "FZEEO.frx":0000
         Left            =   1875
         List            =   "FZEEO.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Tag             =   "Second level of grouping records"
         Top             =   3930
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
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
         Index           =   0
         ItemData        =   "FZEEO.frx":0004
         Left            =   1875
         List            =   "FZEEO.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "First Level of grouping records"
         Top             =   3630
         Width           =   2325
      End
      Begin VB.ComboBox comCountryOfEmp 
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
         Left            =   2100
         TabIndex        =   16
         Tag             =   "00-Country of Employment"
         Top             =   480
         Width           =   1440
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   20
         Tag             =   "00-Enter Location Code"
         Top             =   1200
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         Height          =   285
         Left            =   1800
         TabIndex        =   21
         Top             =   1560
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.DateLookup dlpAsOf 
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Tag             =   "40-From Date"
         Top             =   3045
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1315
      End
      Begin INFOHR_Controls.CodeLookup clpNOGC 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Tag             =   "Enter Job Category Code"
         Top             =   2670
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   6
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   1950
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   23
         Tag             =   "00-Enter Section Code"
         Top             =   2310
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.DateLookup dlpToDate 
         Height          =   285
         Left            =   3840
         TabIndex        =   26
         Tag             =   "40-From Date"
         Top             =   3045
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1315
      End
      Begin VB.Label lblTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3480
         TabIndex        =   39
         Top             =   3090
         Visible         =   0   'False
         Width           =   195
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
         Left            =   120
         TabIndex        =   38
         Top             =   840
         Width           =   510
      End
      Begin VB.Label lblNOC 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOC Code"
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
         TabIndex        =   37
         Top             =   2670
         Width           =   765
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
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblPT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         TabIndex        =   35
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Selection Criteria"
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
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblAsOf 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Term Date From: "
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
         TabIndex        =   33
         Top             =   3045
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #2"
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
         Left            =   120
         TabIndex        =   32
         Top             =   3960
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #1"
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
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   3645
         Width           =   885
      End
      Begin VB.Label lblRepGrp 
         Caption         =   "Report Grouping"
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
         Left            =   0
         TabIndex        =   30
         Top             =   3405
         Width           =   1575
      End
      Begin VB.Label lblCountry 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Employment"
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
         TabIndex        =   29
         Top             =   480
         Width           =   1620
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
         Left            =   120
         TabIndex        =   28
         Top             =   1920
         Width           =   555
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
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   540
      End
   End
   Begin VB.Frame fraPurgeEEO 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   4920
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.ComboBox comCountryOfEPL 
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
         Left            =   2220
         TabIndex        =   2
         Tag             =   "00-Country of Employment"
         Top             =   360
         Width           =   1440
      End
      Begin VB.CommandButton cmdDStart 
         Appearance      =   0  'Flat
         Caption         =   "Update"
         Height          =   330
         Left            =   1320
         TabIndex        =   1
         Tag             =   "Terminate the Employee Selected"
         Top             =   2160
         Width           =   2220
      End
      Begin INFOHR_Controls.DateLookup dlpFDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Tag             =   "40-From Date"
         Top             =   840
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin INFOHR_Controls.DateLookup dlpTDate 
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Tag             =   "40-To Date"
         Top             =   840
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1240
      End
      Begin VB.Label lblCoun2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Country of Employment"
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
         TabIndex        =   8
         Top             =   360
         Width           =   1620
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Range"
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
         TabIndex        =   7
         Top             =   840
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Selection Criteria"
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
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   3480
         TabIndex        =   5
         Top             =   855
         Width           =   315
      End
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   510
      TabIndex        =   10
      Tag             =   "There are no grouping for this report"
      Top             =   600
      Width           =   3435
      _Version        =   65536
      _ExtentX        =   6059
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Work Force Analysis by Job Group"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   11
      Tag             =   "Chose Report Grouping"
      Top             =   960
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Employer Information Report EEO-1"
      ForeColor       =   16711680
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   6780
      Width           =   10980
      _Version        =   65536
      _ExtentX        =   19368
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
         Left            =   3525
         Top             =   150
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
      End
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   2
      Left            =   510
      TabIndex        =   12
      Tag             =   "Chose Report Grouping"
      Top             =   1320
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Detailed EEO-1 Report"
      ForeColor       =   16711680
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
   Begin VB.Label lblSelCri 
      Caption         =   "Selection Reports"
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
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmREEO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xNewFun As Boolean
Dim locPurgeEEO As Boolean

Public Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub OutPut(Destination As Integer)
Dim x%

On Error GoTo PrntErr

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False
'cmdView.Enabled = False
x% = Cri_SetAll()

'Ticket #25649 Franks 06/23/2014
If optGrouping(2).Value Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Me.vbxCrystal.Destination = Destination
Me.vbxCrystal.WindowTitle = txtTitle
MDIMain.Timer1.Enabled = False
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1
vbxCrystal.Reset
MDIMain.Timer1.Enabled = True
'cmdPrint.Enabled = True
'cmdView.Enabled = True
Exit Sub

PrntErr:
'MsgBox "Error Printing - check your Windows Printer setup"
'Ticket #19304
MsgBox Err.Description

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Public Sub cmdPrint_Click()
Call OutPut(1)
End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdView_Click()
Call OutPut(0)
End Sub

Private Sub cmdView_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub




Private Function Cri_SetAll()
Dim x%, LocstrSelCri
Cri_SetAll = False

On Error GoTo modSetCriteria_Err

'Ticket #25649 Franks 06/23/2014
If optGrouping(2).Value Then
    If Len(dlpAsOf.Text) > 0 Or Len(dlpToDate.Text) > 0 Then
        If Not IsDate(dlpAsOf.Text) Then
            MsgBox "If one date was entered then both dates must be entered,"
            Exit Function
        End If
        If Not IsDate(dlpToDate.Text) Then
            MsgBox "If one date was entered then both dates must be entered,"
            Exit Function
        End If
    End If
    Call DetailedEEO1Process
    Exit Function
End If

Screen.MousePointer = HOURGLASS

If optGrouping(0) Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEEO1.rpt"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    End If
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
End If
If optGrouping(1) Then 'Ticket #18790
    glbiOneWhere = False
    glbstrSelCri = ""
    
    'Call glbCri_DeptUN(clpDept.Text)
    Call Cri_CountryOfEmployment
    Call Cri_Code(0)
    Call Cri_Code(1)
    Call Cri_Code(2)
    Call Cri_Div
    Call Cri_PT
    Call Cri_NOC_Code
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    
    x% = Cri_Sorts()

    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEEO2.rpt"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    End If
End If

Cri_SetAll = True
Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Security Time", "Comp Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Function
Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getLocEGroup(comGroup(0).Text)

''Call setRptLabel(Me, 0)
'If comGroup(0) = "(none)" Then
'    grpField$ = a
'    'Exit Function
'End If

Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1


''grpField$ = getLocEGroup(comGroup(1).Text)
''grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
''Me.vbxCrystal.GroupCondition(1) = grpCond$
'
''dscGroup$ = comGroup(1).Text
''If dscGroup$ = "NOC Code" Then
''    dscGroup$ = "ORGANIZATIONAL UNITS"
''End If
''dscGroup$ = "descGroup" & CStr(2) & "= '" & dscGroup$ & "'"
''Me.vbxCrystal.Formulas(10) = dscGroup$

Cri_Sorts = z% ' next section number to format

End Function

Function getLocEGroup(ShowStr As String)
Dim vPosGroup
Select Case ShowStr
    Case lStr("Location"):          getLocEGroup = "{HRTABL.TB_DESC}"
    Case lStr("Region"):            getLocEGroup = "{tblRegion.TB_DESC}"
    Case lStr("Category"):          getLocEGroup = "{tblPT.TB_DESC}"
    'Case "NOC Code":                getLocEGroup = "{HR_OCCUPATION_CLASS.OC_DESCR}"
    Case "Country":                 getLocEGroup = "{HREEO.EO_WORKCOUNTRY}"
    Case "(none)":                  getLocEGroup = "{HREEO.EO_COMPNO}" '"(none)"
End Select
End Function


Private Sub Cri_CountryOfEmployment()
Dim CountryCri As String

If Len(comCountryOfEmp.Text) > 0 Then
    If Not UCase(comCountryOfEmp.Text) = "ALL" Then
        CountryCri = "({HREEO.EO_WORKCOUNTRY} = '" & comCountryOfEmp.Text & "')"
    End If
End If

If Len(CountryCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CountryCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CountryCri
    End If
    glbiOneWhere = True
End If
End Sub
Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREEO.EO_LOC"
    If intIdx% = 1 Then strCd$ = "HREEO.EO_REGION"
    If intIdx% = 2 Then strCd$ = "HREMP.ED_SECTION"
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"

End If

If Len(CodeCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CodeCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
    End If
    glbiOneWhere = True
End If

End Sub
Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
    End If
    glbiOneWhere = True
End If

End Sub
Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREEO.EO_PT}= '" & clpPT.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_NOC_Code()
Dim EECri As String, OneSet%, x%

If Len(clpNOGC.Text) < 1 Then Exit Sub

EECri = "{HREEO.EO_OCC_CAT}= '" & clpNOGC.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub cmdDStart_Click()
Dim SQLQ As String
Dim rsTB As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim xTotNum As Integer
    
    If Len(comCountryOfEPL.Text) = 0 Then
        MsgBox "Country of Employment is required."
        comCountryOfEPL.SetFocus
        Exit Sub
    End If
    If Len(dlpFDate.Text) = 0 Then
        MsgBox "From Date is required."
        dlpFDate.SetFocus
        Exit Sub
    Else
        If Not IsDate(dlpFDate.Text) Then
            MsgBox "Invalid From Date."
            dlpFDate.SetFocus
            Exit Sub
        End If
    End If
    If Len(dlpTDate.Text) = 0 Then
        MsgBox "To Date is required."
        dlpTDate.SetFocus
        Exit Sub
    Else
        If Not IsDate(dlpTDate.Text) Then
            MsgBox "Invalid To Date."
            dlpTDate.SetFocus
            Exit Sub
        End If
    End If
    
    Msg$ = ""
    'If Not chkDeathProc() Then Exit Sub

    Msg$ = Msg$ & "Are you sure you want to Purge Applicants EEO Records?"

    
    Title$ = ("Confirm")
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
    
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    SQLQ = "DELETE FROM HREEO WHERE EO_TYPE = 'A' "
    If Len(comCountryOfEPL.Text) > 0 And Not (UCase(comCountryOfEPL.Text) = "ALL") Then
        SQLQ = SQLQ & "AND EO_WORKCOUNTRY = '" & comCountryOfEPL.Text & "' "
    End If
    SQLQ = SQLQ & "AND EO_DOH >= " & Date_SQL(dlpFDate.Text) & " "
    SQLQ = SQLQ & "AND EO_DOH <= " & Date_SQL(dlpTDate.Text) & " "
    gdbAdoIhr001.Execute SQLQ, xTotNum
    
    MsgBox xTotNum & " record(s) deleted."
        
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMREEO"
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
glbOnTop = "FRMREEO"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Me.Caption = glbFormCaption
If glbFormCaption = "Purge Applicants EEO Records" Then
    locPurgeEEO = True
End If
If glbFormCaption = "EEO Reports" Then
    locPurgeEEO = False
End If

glbOnTop = "FRMREEO"
Call INI_Controls(Me)

'Ticket #18790 - begin
If glbFormCaption = "EEO Reports" Then
    Call optGrouping_Click(0, 1)
    'xNewFun = False
    'If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
    '    xNewFun = True
    'End If
    'xNewFun = True
    'fraWorkForce.Visible = xNewFun
    Call addCountryItems
    lblPT.Caption = lStr("Category")
    lblRegion.Caption = lStr("Region")
    lblLocation.Caption = lStr("Location")
    lblDiv.Caption = lStr("Division")
    lblSection.Caption = lStr("Section")
    Call comGrpLoad
    If glbWFC Then 'Ticket #25649 Franks 06/23/2014
        fraWorkForce.Top = 1680
        fraWorkForce.Left = 480
        fraWorkForce.Width = 7815
        'optGrouping(2).Visible = True
    Else
        fraWorkForce.Top = 1320
        fraWorkForce.Left = 480
        fraWorkForce.Width = 7815
        optGrouping(2).Visible = False
    End If
End If
If glbFormCaption = "Purge Applicants EEO Records" Then
    Call addCountryItems
    fraPurgeEEO.Top = 120
    fraPurgeEEO.Left = 510
    fraPurgeEEO.Visible = True
    lblSelCri.Visible = False
    cmdDStart.Enabled = gSec_Upd_AffirmAction_Purge
End If
'Ticket #18790 - end

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optGrouping_Click(Index As Integer, Value As Integer)
txtTitle = optGrouping(Index).Caption
'If optGrouping(1).Value Then
'    fraWorkForce.Visible = True
'Else
'    fraWorkForce.Visible = False
'End If
'Ticket #25649 Franks 06/23/2014
If optGrouping(0).Value Then
    fraWorkForce.Visible = False
End If
If optGrouping(1).Value Then
    fraWorkForce.Visible = True
    Call GrpInfoDisp(Index)
End If
If optGrouping(2).Value Then
    fraWorkForce.Visible = True
    Call GrpInfoDisp(Index)
End If

End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub


Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = False
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = IIf(locPurgeEEO, False, True)
End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub


Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub addCountryItems()
Dim ctylist, x

ctylist = CountryList
x = 1
Do While x > 0
    x = InStr(ctylist, "&")
    If x > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, x - 1)
        comCountryOfEPL.AddItem Left(ctylist, x - 1)
        ctylist = Mid(ctylist, x + 1)
    Else
        comCountryOfEmp.AddItem ctylist
        comCountryOfEPL.AddItem ctylist
    End If
Loop

End Sub
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
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comCountryOfEmp) = 0 And comCountryOfEmp <> "" Then
    xCountryList = xCountryList & "&" & comCountryOfEmp
    comCountryOfEmp.AddItem comCountryOfEmp
    comCountryOfEPL.AddItem comCountryOfEPL
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
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Sub comGrpLoad()
Dim vPosGroup As String
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(0).AddItem "Country"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Category")
    'comGroup(0).AddItem "NOC Code"
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 4
    'comGroup(1).AddItem lStr("Region")
    'comGroup(1).AddItem lStr("Location")
    'comGroup(1).AddItem lStr("Category")
    'comGroup(1).AddItem "NOC Code"
    'comGroup(1).ListIndex = 3
End Sub

Private Sub GrpInfoDisp(xInd) 'Ticket #25649 Franks 06/23/2014
    If xInd = 1 Then
        lblRepGrp.Visible = True
        lblGrp(0).Visible = True
        comGroup(0).Visible = True
        lblAsOf.Visible = False
        dlpAsOf.Visible = False
        lblTo.Visible = False
        dlpToDate.Visible = False
    End If
    If xInd = 2 Then
        lblRepGrp.Visible = False
        lblGrp(0).Visible = False
        comGroup(0).Visible = False
        lblAsOf.Visible = True
        dlpAsOf.Visible = True
        lblTo.Visible = True
        dlpToDate.Visible = True
    End If
End Sub

Private Sub DetailedEEO1Process() 'Ticket #25649 Franks 06/23/2014
Dim rsEEO As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim SQLQ As String
Dim I As Long
Dim xTot As Long
Dim xEmpNo
Dim xNOC
Dim xNoRecAct As Boolean
Dim xNoRecTerm As Boolean

    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ

    'For active employees - begin ---------------------------
    SQLQ = "SELECT HREEO.*,ED_SECTION,ED_DIV,ED_SIN,ED_SEX,ED_DEPTNO,ED_CITY,ED_PROV,ED_PCODE,ED_DOH FROM HREEO LEFT JOIN HREMP ON HREEO.EO_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE (1=1) "
    SQLQ = SQLQ & "AND " & glbSeleDeptUn & " "
    If Len(comCountryOfEmp.Text) > 0 Then
        If Not UCase(comCountryOfEmp.Text) = "ALL" Then
            SQLQ = SQLQ & "AND EO_WORKCOUNTRY = '" & comCountryOfEmp.Text & "' "
        End If
    End If
    If Len(clpCode(0).Text) > 0 Then
        SQLQ = SQLQ & "AND EO_LOC = '" & clpCode(0).Text & "' "
    End If
    If Len(clpCode(1).Text) > 0 Then
        SQLQ = SQLQ & "AND EO_REGION = '" & clpCode(1).Text & "' "
    End If
    If Len(clpCode(2).Text) > 0 Then
        SQLQ = SQLQ & "AND ED_SECTION = '" & clpCode(2).Text & "' "
    End If
    If Len(clpPT.Text) > 0 Then
        SQLQ = SQLQ & "AND EO_PT = '" & clpPT.Text & "' "
    End If
    If Len(clpDiv.Text) > 0 Then
        SQLQ = SQLQ & "AND ED_DIV = '" & clpDiv.Text & "' "
    End If
    If Len(clpNOGC.Text) > 0 Then
        SQLQ = SQLQ & "AND EO_OCC_CAT = '" & clpNOGC.Text & "' "
    End If
    If glbNoNONE Then
        SQLQ = SQLQ & "AND NOT HREMP.ED_ORG = 'NONE' "
    End If
    If glbNoEXEC Then
        SQLQ = SQLQ & "AND NOT HREMP.ED_ORG = 'EXEC' "
    End If
    
    xNoRecAct = False
    If rsEEO.State <> 0 Then rsEEO.Close
    rsEEO.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEEO.EOF Then
        xNoRecAct = True
        'MsgBox "No record found in this Selection Criteria."
        'Exit Sub
    End If
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    If Not rsEEO.EOF Then
        xTot = rsEEO.RecordCount
    End If
    I = 0
    I = 0
    Do While Not rsEEO.EOF
        MDIMain.panHelp(0).FloodPercent = Int((I / xTot) * 100)
        I = I + 1
        DoEvents
        xEmpNo = rsEEO("EO_EMPNBR")
        
        'add to the wrk table
        SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        SQLQ = SQLQ & "AND TT_EMPNBR = " & xEmpNo & " "
        If rsWRK.State <> 0 Then rsWRK.Close
        rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWRK.EOF Then
            rsWRK.AddNew
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK("TT_EMPNBR") = xEmpNo
        End If
        rsWRK("TT_CHAR10") = rsEEO("ED_DIV")
        rsWRK("TT_NAMEFLD") = Left(rsEEO("EO_SURNAME") & ", " & rsEEO("EO_FNAME"), 40)
        If Not IsNull(rsEEO("ED_SIN")) Then rsWRK("TT_RESULTD") = Left(rsEEO("ED_SIN"), 10)
        If Not IsNull(rsEEO("EO_RACE")) Then rsWRK("TT_SHIFT") = Left(rsEEO("EO_RACE"), 4)
        rsWRK("TT_SEX") = rsEEO("ED_SEX")
        If IsNull(rsEEO("EO_DISABLE_YN")) Then
            rsWRK("TT_SPC1") = "N"
        Else
            If rsEEO("EO_DISABLE_YN") Then rsWRK("TT_SPC1") = "Y" Else rsWRK("TT_SPC1") = "N"
        End If
        If IsNull(rsEEO("EO_VETERAN")) Then
            rsWRK("TT_SPC2") = "N"
        Else
            If rsEEO("EO_VETERAN") Then rsWRK("TT_SPC2") = "Y" Else rsWRK("TT_SPC2") = "N"
        End If
        If IsNull(rsEEO("EO_VIETNAM")) Then
            rsWRK("TT_SPC3") = "N"
        Else
            If rsEEO("EO_VIETNAM") Then rsWRK("TT_SPC3") = "Y" Else rsWRK("TT_SPC3") = "N"
        End If
        'Salary info - begin
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & xEmpNo & " "
        If rsSal.State <> 0 Then rsSal.Close
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsSal.EOF Then
            If Not IsNull(rsSal("SH_SREAS1")) Then rsWRK("TT_OLDDIV") = rsSal("SH_SREAS1")
            If Not IsNull(rsSal("SH_SALARY")) Then rsWRK("TT_SALARY") = rsSal("SH_SALARY")
            If IsDate(rsSal("SH_EDATE")) Then rsWRK("TT_SEDATE") = rsSal("SH_EDATE") '
            If Not IsNull(rsSal("SH_SALCD")) Then rsWRK("TT_SALCD") = rsSal("SH_SALCD")
            If Not IsNull(rsSal("SH_BAND")) Then rsWRK("TT_COMPL") = rsSal("SH_BAND")
            rsWRK("TT_JOB") = Left(GetJobFieldVal("JB_DESCR", rsSal("SH_JOB")), 50)
            rsWRK("TT_SDATE") = rsSal("SH_SDATE")
            rsWRK("TT_SCHOOL") = Left(GetJobFieldVal("JB_FEDGRP", rsSal("SH_JOB")), 50) 'NOC
        End If '
        rsSal.Close
        'Salary info - end
        rsWRK("TT_NEWDEPT") = rsEEO("ED_DEPTNO")
        rsWRK("TT_LANG1") = rsEEO("ED_CITY")
        If Not IsNull(rsEEO("ED_PROV")) Then rsWRK("TT_GRADE") = Left(rsEEO("ED_PROV"), 2)
        If Not IsNull(rsEEO("ED_PCODE")) Then rsWRK("TT_PREPT") = Left(rsEEO("ED_PCODE"), 10)
        rsWRK("TT_DATEFLD") = rsEEO("ED_DOH")
        rsWRK.Update
        rsEEO.MoveNext
    Loop
    rsEEO.Close
    'For active employees - end ---------------------------
    
    'For term employees - begin ---------------------------
    xNoRecTerm = True
    If IsDate(dlpAsOf.Text) And IsDate(dlpToDate.Text) Then
        SQLQ = "SELECT Term_HREEO.*,ED_SECTION,ED_DIV,ED_SIN,ED_SEX,ED_DEPTNO,ED_CITY,ED_PROV,ED_PCODE,ED_DOH FROM Term_HREEO LEFT JOIN Term_HREMP ON Term_HREEO.TERM_SEQ = Term_HREMP.TERM_SEQ "
        SQLQ = SQLQ & "LEFT JOIN TERM_HRTRMEMP ON Term_HREMP.TERM_SEQ = TERM_HRTRMEMP.TERM_SEQ "
        SQLQ = SQLQ & "WHERE (1=1) "
        SQLQ = SQLQ & "AND " & glbSeleDeptUn & " "
        SQLQ = SQLQ & "AND TERM_HRTRMEMP.Term_DOT >= " & Date_SQL(dlpAsOf.Text) & " "
        SQLQ = SQLQ & "AND TERM_HRTRMEMP.Term_DOT <= " & Date_SQL(dlpToDate.Text) & " "
        If Len(comCountryOfEmp.Text) > 0 Then
            If Not UCase(comCountryOfEmp.Text) = "ALL" Then
                SQLQ = SQLQ & "AND EO_WORKCOUNTRY = '" & comCountryOfEmp.Text & "' "
            End If
        End If
        If Len(clpCode(0).Text) > 0 Then
            SQLQ = SQLQ & "AND EO_LOC = '" & clpCode(0).Text & "' "
        End If
        If Len(clpCode(1).Text) > 0 Then
            SQLQ = SQLQ & "AND EO_REGION = '" & clpCode(1).Text & "' "
        End If
        If Len(clpCode(2).Text) > 0 Then
            SQLQ = SQLQ & "AND ED_SECTION = '" & clpCode(2).Text & "' "
        End If
        If Len(clpPT.Text) > 0 Then
            SQLQ = SQLQ & "AND EO_PT = '" & clpPT.Text & "' "
        End If
        If Len(clpDiv.Text) > 0 Then
            SQLQ = SQLQ & "AND ED_DIV = '" & clpDiv.Text & "' "
        End If
        If Len(clpNOGC.Text) > 0 Then
            SQLQ = SQLQ & "AND EO_OCC_CAT = '" & clpNOGC.Text & "' "
        End If
        If glbNoNONE Then
            SQLQ = SQLQ & "AND NOT Term_HREMP.ED_ORG = 'NONE' "
        End If
        If glbNoEXEC Then
            SQLQ = SQLQ & "AND NOT Term_HREMP.ED_ORG = 'EXEC' "
        End If

        If rsEEO.State <> 0 Then rsEEO.Close
        rsEEO.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xNoRecTerm = False
        If rsEEO.EOF Then
            xNoRecTerm = True
            'MsgBox "No record found in this Selection Criteria."
            'Exit Sub
        End If
        Screen.MousePointer = HOURGLASS
        MDIMain.panHelp(0).FloodType = 1
        If Not rsEEO.EOF Then
            xTot = rsEEO.RecordCount
        End If
        I = 0
        I = 0
        Do While Not rsEEO.EOF
            MDIMain.panHelp(0).FloodPercent = Int((I / xTot) * 100)
            I = I + 1
            DoEvents
            xEmpNo = rsEEO("EO_EMPNBR")
            
            'add to the wrk table
            SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
            SQLQ = SQLQ & "AND TT_EMPNBR = " & xEmpNo & " "
            If rsWRK.State <> 0 Then rsWRK.Close
            rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsWRK.EOF Then
                rsWRK.AddNew
                rsWRK("TT_WRKEMP") = glbUserID
                rsWRK("TT_EMPNBR") = xEmpNo
            End If
            rsWRK("TT_CHAR10") = rsEEO("ED_DIV")
            rsWRK("TT_NAMEFLD") = Left(rsEEO("EO_SURNAME") & ", " & rsEEO("EO_FNAME"), 40)
            If Not IsNull(rsEEO("ED_SIN")) Then rsWRK("TT_RESULTD") = Left(rsEEO("ED_SIN"), 10)
            If Not IsNull(rsEEO("EO_RACE")) Then rsWRK("TT_SHIFT") = Left(rsEEO("EO_RACE"), 4)
            rsWRK("TT_SEX") = rsEEO("ED_SEX")
            If IsNull(rsEEO("EO_DISABLE_YN")) Then
                rsWRK("TT_SPC1") = "N"
            Else
                If rsEEO("EO_DISABLE_YN") Then rsWRK("TT_SPC1") = "Y" Else rsWRK("TT_SPC1") = "N"
            End If
            If IsNull(rsEEO("EO_VETERAN")) Then
                rsWRK("TT_SPC2") = "N"
            Else
                If rsEEO("EO_VETERAN") Then rsWRK("TT_SPC2") = "Y" Else rsWRK("TT_SPC2") = "N"
            End If
            If IsNull(rsEEO("EO_VIETNAM")) Then
                rsWRK("TT_SPC3") = "N"
            Else
                If rsEEO("EO_VIETNAM") Then rsWRK("TT_SPC3") = "Y" Else rsWRK("TT_SPC3") = "N"
            End If
            'Salary info - begin
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & xEmpNo & " "
            If rsSal.State <> 0 Then rsSal.Close
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsSal.EOF Then
                If Not IsNull(rsSal("SH_SREAS1")) Then rsWRK("TT_OLDDIV") = rsSal("SH_SREAS1")
                If Not IsNull(rsSal("SH_SALARY")) Then rsWRK("TT_SALARY") = rsSal("SH_SALARY")
                If IsDate(rsSal("SH_EDATE")) Then rsWRK("TT_SEDATE") = rsSal("SH_EDATE") '
                If Not IsNull(rsSal("SH_SALCD")) Then rsWRK("TT_SALCD") = rsSal("SH_SALCD")
                If Not IsNull(rsSal("SH_BAND")) Then rsWRK("TT_COMPL") = rsSal("SH_BAND")
                rsWRK("TT_JOB") = Left(GetJobFieldVal("JB_DESCR", rsSal("SH_JOB")), 50)
                rsWRK("TT_SDATE") = rsSal("SH_SDATE")
                rsWRK("TT_SCHOOL") = Left(GetJobFieldVal("JB_FEDGRP", rsSal("SH_JOB")), 50) 'NOC
            End If '
            rsSal.Close
            'Salary info - end
            rsWRK("TT_NEWDEPT") = rsEEO("ED_DEPTNO")
            rsWRK("TT_LANG1") = rsEEO("ED_CITY")
            If Not IsNull(rsEEO("ED_PROV")) Then rsWRK("TT_GRADE") = Left(rsEEO("ED_PROV"), 2)
            If Not IsNull(rsEEO("ED_PCODE")) Then rsWRK("TT_PREPT") = Left(rsEEO("ED_PCODE"), 10)
            rsWRK("TT_DATEFLD") = rsEEO("ED_DOH")
            rsWRK("TT_NUMERIC") = rsEEO("TERM_SEQ")
            rsWRK.Update '
            rsEEO.MoveNext
        Loop
        rsEEO.Close
    End If
    'For term employees - end ---------------------------
    
    If xNoRecAct And xNoRecTerm Then
        MsgBox "No record found in this Selection Criteria."
        Exit Sub
    End If
    
    Call XLSwriter_EEO1
    
End Sub

Private Function GetJobFieldVal(xField, xCode) 'Ticket #23876 Franks 06/10/2013
Dim rsJOB As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsJOB.EOF Then
            xRetVal = rsJOB(xField) ' rsJOB("JB_DESCR")
        End If
        rsJOB.Close
    End If
    GetJobFieldVal = xRetVal
End Function

Private Sub XLSwriter_EEO1()
Dim SQLQ As String
'Dim exApp As Excel.Application, exBook As Excel.Workbook, exSheet As Excel.Worksheet
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xlsFileTmp As String, xlsFileMat As String
Dim rsIn As New ADODB.Recordset, rsOUT As New ADODB.Recordset, rsEmp As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsCurJob As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim xRow As Long, xCol As Long, xwCol As Long
Dim xType As String, strTemp As String, strDate As String
Dim NewDateFormat As String, flgReqC As Boolean, strDisp As String, xMax As Long, retval As Long, I As Integer
Dim xStartDate As String
Dim xTrainMatrixPath

On Error GoTo Err_XLS
    
    xTrainMatrixPath = ""
    If gsTRAININGMATRIX Then
        xTrainMatrixPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xTrainMatrixPath) = 0 Then
        xTrainMatrixPath = glbIHRREPORTS
    End If

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "WFC_DetailedEEOTmp.xls"
    xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "WFC_DetailedEEO(" & Trim(glbUserID) & ").xls"

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
'        'Create new WorkBook of Excel
'    Set exApp = New Excel.Application
'    Set exBook = exApp.Workbooks.Open(xlsFileMat)

    
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "ORDER BY TT_CHAR10,TT_NAMEFLD "
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWRK.EOF Then
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)

        exSheet.Cells(2, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
        exSheet.Cells(3, 1) = "Time: " & Time$
                
        xRow = 5
        I = 0
        xMax = rsWRK.RecordCount
        Do While Not rsWRK.EOF
            MDIMain.panHelp(0).FloodPercent = (I / xMax) * 100
            I = I + 1
            DoEvents
            
            exSheet.Cells(xRow, 1) = rsWRK("TT_CHAR10")
            exSheet.Cells(xRow, 2) = rsWRK("TT_NAMEFLD")
            exSheet.Cells(xRow, 3) = rsWRK("TT_RESULTD")
            exSheet.Cells(xRow, 4) = rsWRK("TT_SHIFT")
            exSheet.Cells(xRow, 5) = rsWRK("TT_SEX")
            exSheet.Cells(xRow, 6) = rsWRK("TT_SPC1")
            exSheet.Cells(xRow, 7) = rsWRK("TT_SPC2")
            exSheet.Cells(xRow, 8) = rsWRK("TT_SPC3")
            exSheet.Cells(xRow, 9) = rsWRK("TT_OLDDIV")
            exSheet.Cells(xRow, 10) = rsWRK("TT_SALARY")
            exSheet.Cells(xRow, 11) = rsWRK("TT_SEDATE")
            exSheet.Cells(xRow, 12) = rsWRK("TT_SALCD")
            exSheet.Cells(xRow, 13) = rsWRK("TT_NEWDEPT")
            exSheet.Cells(xRow, 14) = rsWRK("TT_LANG1") 'city
            exSheet.Cells(xRow, 15) = rsWRK("TT_GRADE") 'state
            exSheet.Cells(xRow, 16) = rsWRK("TT_PREPT") 'zip
            exSheet.Cells(xRow, 17) = rsWRK("TT_DATEFLD") 'doh
            exSheet.Cells(xRow, 18) = rsWRK("TT_COMPL") 'band
            exSheet.Cells(xRow, 19) = rsWRK("TT_JOB")
            exSheet.Cells(xRow, 20) = rsWRK("TT_SDATE")
            exSheet.Cells(xRow, 21) = rsWRK("TT_SCHOOL")
            rsWRK.MoveNext
            xRow = xRow + 1
        Loop
        rsWRK.Close
    
        'Save new Excel file as XLS
        'exBook.SaveAs "C:\TrainMat.xls"
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(1).Caption = ""
        MDIMain.panHelp(2).Caption = ""
    
        Call Pause(1)
        'launch Excel file
        'Shell "Start " & GetShortName(xlsFileMat)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
        
        Exit Sub
    
    End If
exH:
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
    Exit Sub
Err_XLS:

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Detailed EEO-1 Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Sub
    End If
    If Not exApp Is Nothing Then
        If exApp.Visible = False Then
            exApp.Quit
        End If
        Set exApp = Nothing
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter_EEO1", "", "Select")
'Resume Next
End Sub

