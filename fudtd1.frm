VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmUTd1Dollar 
   Appearance      =   0  'Flat
   Caption         =   "TD1 Dollar"
   ClientHeight    =   6255
   ClientLeft      =   1560
   ClientTop       =   2325
   ClientWidth     =   9480
   DrawMode        =   1  'Blackness
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   9480
   Tag             =   "Dollar Entitlements Mass Update"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtProvNCode 
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
      Left            =   5340
      MaxLength       =   2
      TabIndex        =   19
      Tag             =   "10-Enter New TD1 Code"
      Top             =   5160
      Width           =   990
   End
   Begin VB.TextBox txtProvOCode 
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
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   18
      Tag             =   "10-Enter Old TD1 Code"
      Top             =   5160
      Width           =   990
   End
   Begin VB.TextBox txtProvNDollar 
      Appearance      =   0  'Flat
      DataField       =   "ED_PROVAMT"
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
      Left            =   5340
      MaxLength       =   6
      TabIndex        =   16
      Tag             =   "11-Enter New TD1 Dollar"
      Top             =   4830
      Width           =   990
   End
   Begin VB.TextBox txtProvODollar 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   15
      Tag             =   "11-Enter Old TD1 Dollar"
      Top             =   4830
      Width           =   990
   End
   Begin VB.TextBox txtNCode 
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
      Left            =   5340
      MaxLength       =   4
      TabIndex        =   14
      Tag             =   "10-Enter New TD1 Code"
      Top             =   4350
      Width           =   990
   End
   Begin VB.TextBox txtOCode 
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
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "10-Enter Old TD1 Code"
      Top             =   4350
      Width           =   990
   End
   Begin VB.TextBox txtNDollar 
      Appearance      =   0  'Flat
      DataField       =   "ED_TD1DOL"
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
      Left            =   5340
      MaxLength       =   6
      TabIndex        =   11
      Tag             =   "11-Enter New TD1 Dollar"
      Top             =   4020
      Width           =   990
   End
   Begin VB.TextBox txtODollar 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   10
      Tag             =   "11-Enter Old TD1 Dollar"
      Top             =   4020
      Width           =   990
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3570
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3240
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2910
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2250
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1920
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1590
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1260
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   930
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2580
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6835
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSCheck chkBasicTD1Amt 
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Tag             =   "To update with Basic TD1 Dollar from Company Master independent of Old TD1 Dollar"
      Top             =   4035
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Update with Basic TD1 Dollar"
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
   Begin Threed.SSCheck chkBasicProvAmt 
      Height          =   255
      Left            =   6600
      TabIndex        =   17
      Tag             =   "To update with Basic Prov. Dollar from Company Master independent of Old Prov. Dollar"
      Top             =   4845
      Width           =   2535
      _Version        =   65536
      _ExtentX        =   4471
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Update with Basic Prov. Dollar"
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
      Left            =   420
      TabIndex        =   38
      Top             =   1250
      Width           =   615
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
      Left            =   420
      TabIndex        =   37
      Top             =   570
      Width           =   555
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   420
      TabIndex        =   36
      Top             =   930
      Width           =   825
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
      Left            =   420
      TabIndex        =   35
      Top             =   1590
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      Left            =   420
      TabIndex        =   34
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      Left            =   420
      TabIndex        =   33
      Top             =   2580
      Width           =   1290
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   420
      TabIndex        =   32
      Top             =   3240
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   420
      TabIndex        =   31
      Top             =   2910
      Width           =   510
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
      Left            =   420
      TabIndex        =   30
      Top             =   3540
      Width           =   540
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
      Left            =   420
      TabIndex        =   29
      Top             =   2250
      Width           =   630
   End
   Begin VB.Label lblProvNCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Prov. Code:"
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
      Left            =   3780
      TabIndex        =   28
      Top             =   5190
      Width           =   1560
   End
   Begin VB.Label lblProvOCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Prov. Code:"
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
      Top             =   5190
      Width           =   1560
   End
   Begin VB.Label lblProvNDollar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New Prov. Dollar:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3780
      TabIndex        =   26
      Top             =   4830
      Width           =   1560
   End
   Begin VB.Label lblProvODollar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Prov. Dollar:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   4860
      Width           =   1560
   End
   Begin VB.Label lblNCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New TD1 Code:"
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
      Left            =   3780
      TabIndex        =   24
      Top             =   4380
      Width           =   1560
   End
   Begin VB.Label lblOCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Old TD1 Code:"
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
      TabIndex        =   23
      Top             =   4380
      Width           =   1560
   End
   Begin VB.Label lblNDollar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "New TD1 Dollar:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3780
      TabIndex        =   22
      Top             =   4050
      Width           =   1560
   End
   Begin VB.Label lblODollar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Old TD1 Dollar:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   4050
      Width           =   1560
   End
   Begin VB.Label lblSelCri 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   20
      Top             =   270
      Width           =   1575
   End
End
Attribute VB_Name = "frmUTd1Dollar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbDelete%

Private Function chkTD1() As Boolean
Dim X%
chkTD1 = False

If Len(clpDiv) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
    clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
    clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Len(clpCode(X%)) > 0 And clpCode(X%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X%).SetFocus
        Exit Function
    End If
Next X%

If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpPT.SetFocus
    Exit Function
End If

If Not chkBasicTD1Amt And Not chkBasicProvAmt Then
    ' Check to make sure they're actually trying to change something
    If Len(txtODollar.Text) = 0 And Len(txtProvODollar.Text) = 0 Then
        MsgBox "Old TD1 Dollar and/or Old Prov. Dollar must be entered."
        txtODollar.SetFocus
        Exit Function
    End If
    ' If they want to change TD1$ they need both old and new
    If Len(txtODollar.Text) > 0 And Len(txtNDollar.Text) = 0 Then
        MsgBox "If Old TD1 Dollar is entered, New TD1 Dollar must be entered."
        txtNDollar.SetFocus
        Exit Function
    End If
    ' If they want to change TD1 Code, they need TD1$ too.
    If txtOCode <> "" And txtODollar = "" Then
        MsgBox "If Old or New TD1 Code is entered, Old & New TD1 Dollar must also be entered."
        txtNDollar.SetFocus
        Exit Function
    End If
    
    ' If they want to change Prov$ they need both old and new
    If Len(txtProvODollar.Text) > 0 And Len(txtProvNDollar.Text) = 0 Then
        MsgBox "If Old Prov. Dollar is entered, New Prov. Dollar must be entered."
        txtProvNDollar.SetFocus
        Exit Function
    End If
    ' If they want to change Prov. Code, they need Prov$ too.
    If txtProvOCode <> "" And txtProvODollar = "" Then
        MsgBox "If Old or New Prov. Code is entered, Old & New Prov. Dollar must also be entered."
        txtNDollar.SetFocus
        Exit Function
    End If
End If

If chkBasicTD1Amt Then
    If Len(txtNDollar.Text) = 0 Then
        MsgBox "New TD1 Dollar cannot be blank. Please update Setup\Company Master screen with Basic TD1 Dollar."
        Exit Function
    End If
End If

If chkBasicProvAmt Then
    If Len(txtProvNDollar.Text) = 0 Then
        MsgBox "New Prov. Dollar cannot be blank. Please update Setup\Company Master screen with Basic Prov. Dollar."
        Exit Function
    End If
End If

'Ticket #23557 - (New condition) Added this to accomodate the next condition changed by Frank.
'None of the values entered - nothing to update
If (Not chkBasicTD1Amt And Len(txtNDollar.Text) = 0 And Len(txtODollar.Text) = 0) And (Not chkBasicProvAmt And Len(txtProvNDollar.Text) = 0 And Len(txtProvODollar.Text) = 0) Then
    MsgBox "No values entered for TD1 Dollar/Code or Prov. Dollar/Code to update."
    txtNDollar.SetFocus
    Exit Function
End If

'Ticket #23557 - Added new condition, Frank's modified condition not working for Jerry
If Not chkBasicTD1Amt And ((Len(txtNDollar.Text) > 0 And Len(txtODollar.Text) = 0) Or (Len(txtNDollar.Text) = 0 And Len(txtODollar.Text) > 0)) Then
'Ticket #21351 Franks 12/23/2011
'If Not chkBasicTD1Amt And (Len(txtNDollar.Text) = 0 Or Len(txtODollar.Text) = 0) Then
    MsgBox "Both, Old TD1 Dollar and New TD1 Dollar, must be entered."
    txtNDollar.SetFocus
    Exit Function
End If

'Ticket #23557 - Added new condition, Frank's modified condition not working for Jerry
If Not chkBasicProvAmt And ((Len(txtProvNDollar.Text) > 0 And Len(txtProvODollar.Text) = 0) Or (Len(txtProvNDollar.Text) = 0 And Len(txtProvODollar.Text) > 0)) Then
'If Not chkBasicProvAmt And (Len(txtProvNDollar.Text) > 0 Or Len(txtProvODollar.Text) > 0) Then
'Ticket #21351 Franks 12/23/2011
'If Not chkBasicProvAmt And (Len(txtProvNDollar.Text) = 0 Or Len(txtProvODollar.Text) = 0) Then
    MsgBox "Both, Old Prov. Dollar and New Prov. Dollar, must be entered."
    txtNDollar.SetFocus
    Exit Function
End If

chkTD1 = True
End Function

Public Sub cmdClose_Click()
Unload Me
End Sub


Public Sub cmdModify_Click()
Dim Msg$, DgDef As Variant, Response%
Dim dd&
Dim lngRecs&
Dim pct%, prec%
Dim SQLQ, SQLQW, WSQLQ, ESQLQ, rsEmp As New ADODB.Recordset
Dim flgTD1Upd, flgProvUpd As Boolean

If Not chkTD1() Then Exit Sub

flgTD1Upd = False
flgProvUpd = False

WSQLQ = " WHERE " & glbSeleDeptUn
If Len(clpDept) > 0 Then WSQLQ = WSQLQ & " AND ED_DEPTNO = '" & clpDept & "'"
If Len(clpDiv) > 0 Then WSQLQ = WSQLQ & " AND ED_DIV = '" & clpDiv & "' "
If Len(clpCode(1).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ORG = '" & clpCode(1).Text & "' "
If Len(clpCode(2).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMP = '" & clpCode(2).Text & "' "
If Len(clpPT) > 0 Then WSQLQ = WSQLQ & " AND ED_PT = '" & clpPT & "' "
If Len(clpCode(3)) > 0 Then
    If glbLinamar Then
         WSQLQ = WSQLQ & " AND (ED_REGION = '" & clpDiv & clpCode(3) & "' or  ED_REGION= 'ALL" & clpCode(3) & "')"
    Else
         WSQLQ = WSQLQ & " AND ED_REGION = '" & clpCode(3) & "' "
    End If
End If
If Len(clpCode(4).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ADMINBY = '" & clpCode(4).Text & "' "
If Len(clpCode(5).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_SECTION = '" & clpCode(5).Text & "' "
If Len(elpEEID.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID) & ") "

ESQLQ = "SELECT ED_EMPNBR FROM HREMP " & WSQLQ
rsEmp.Open ESQLQ, gdbAdoIhr001, adOpenKeyset

If rsEmp.EOF And rsEmp.BOF Then
  MsgBox "Records for this selection do not exist!"
  Screen.MousePointer = DEFAULT
  Exit Sub
End If

rsEmp.Close

' Change the TD1$ (and TD1 Code if requested)
If Len(txtODollar) > 0 Or chkBasicTD1Amt Then
    SQLQW = WSQLQ
    If Len(txtODollar) > 0 Then SQLQW = SQLQW & " AND ED_TD1DOL = " & txtODollar.Text
    If Len(txtOCode.Text) > 0 Then SQLQW = SQLQW & " AND ED_TD1CODE = '" & txtOCode.Text & "' "
    
    ESQLQ = "SELECT ED_EMPNBR FROM HREMP " & SQLQW  ' WSQLQ
    rsEmp.Open ESQLQ, gdbAdoIhr001, adOpenKeyset
    prec% = 0
    lngRecs& = rsEmp.RecordCount
        
    If lngRecs& > 0 Then
        Msg$ = lngRecs& & " TD1 Dollar/Code Records to process" & Chr(10) & "Do you want to Proceed?"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, "TD1 Dollar/Code Update")    ' Get user response.
        If Response% = IDNO Then ' Evaluate response
            rsEmp.Close
            Set rsEmp = Nothing
            flgTD1Upd = False

            GoTo ProvAmt_Update
        End If
    Else
        MsgBox "No TD1 Dollar/Code record found to process for this selection criteria."
        rsEmp.Close
        Set rsEmp = Nothing
        flgTD1Upd = False

        GoTo ProvAmt_Update
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodShowPct = True
    
    Do Until rsEmp.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        Call Transfer_TaxAmount(rsEmp("ED_EMPNBR").Value, txtODollar.Text, txtNDollar)
        rsEmp.MoveNext
        DoEvents
    Loop
    rsEmp.Close
    
    SQLQ = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_TD1DOL,AU_OLDTD1,"
    SQLQ = SQLQ & "AU_PAYROLL_ID," 'If glbSoroc Then
    SQLQ = SQLQ & "AU_TD1CODE,AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
    If Not glbSQL Then SQLQ = SQLQ & " IN '" & glbIHRAUDIT & "' [;PWD=petman;DATABASE=" & glbIHRAUDIT & "] "
    
    SQLQ = SQLQ & " SELECT 'M','N',ED_EMPNBR, " & txtNDollar & "," & IIf(Len(txtODollar) = 0 And chkBasicTD1Amt, "ED_TD1DOL", txtODollar)
    SQLQ = SQLQ & ",ED_PAYROLL_ID" 'If glbSoroc Then
    SQLQ = SQLQ & "," & IIf(Len(txtNCode) > 0, "ED_TD1CODE", "Null")
    SQLQ = SQLQ & ",ED_DIV,ED_PT, "
    SQLQ = SQLQ & Date_SQL(Now) & " As AU_LDATE, '"
    SQLQ = SQLQ & Time$ & "' As AU_LTIME,'N','" & glbUserID & "' As AU_LUSER FROM HREMP "
    gdbAdoIhr001.Execute SQLQ & SQLQW
    
    SQLQ = "UPDATE HREMP SET ED_TD1DOL = " & txtNDollar
    If Len(txtNCode.Text) > 0 Then SQLQ = SQLQ & ",ED_TD1CODE = '" & txtNCode.Text & "' "
    gdbAdoIhr001.Execute SQLQ & SQLQW
    
    Screen.MousePointer = DEFAULT
    
    flgTD1Upd = True
End If

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).FloodShowPct = False

ProvAmt_Update:
' Change the Prov$ (and Prov. Code if requested)
If Len(txtProvODollar) > 0 Or chkBasicProvAmt Then
    SQLQW = WSQLQ
    If Len(txtProvODollar) > 0 Then SQLQW = SQLQW & " AND ED_PROVAMT = " & txtProvODollar.Text
    If Len(txtProvOCode.Text) > 0 Then SQLQW = SQLQW & " AND ED_PROVCODE = '" & txtProvOCode & "' "
    
    ESQLQ = "SELECT ED_EMPNBR FROM HREMP " & SQLQW  ' WSQLQ
    rsEmp.Open ESQLQ, gdbAdoIhr001, adOpenKeyset
    prec% = 0
    lngRecs& = rsEmp.RecordCount
    
    If lngRecs& > 0 Then
        Msg$ = lngRecs& & " Provincial Dollar/Code Records to process" & Chr(10) & "Do you want to Proceed?"
        DgDef = MB_YESNO + MB_ICONEXCLAMATION + MB_DEFBUTTON2  ' Describe dialog.
        Response% = MsgBox(Msg, DgDef, "Provincial Dollar/Code Update")    ' Get user response.
        If Response% = IDNO Then ' Evaluate response
            rsEmp.Close
            Set rsEmp = Nothing
            flgProvUpd = False
            
            GoTo Continue
        End If
    Else
        MsgBox "No Provincial Dollar/Code record found to process for this selection criteria."
        rsEmp.Close
        Set rsEmp = Nothing
        flgProvUpd = False

        GoTo Continue
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).FloodShowPct = True
    MDIMain.panHelp(0).FloodType = 1
    DoEvents
    
    Do Until rsEmp.EOF
        prec% = prec% + 1
        pct% = Int(100 * (prec% / (lngRecs&)))
        MDIMain.panHelp(0).FloodPercent = pct%
        Call Transfer_TaxAmount(rsEmp("ED_EMPNBR").Value, txtProvODollar.Text, txtProvNDollar)
        rsEmp.MoveNext
        DoEvents
    Loop
    rsEmp.Close
    
    
    SQLQ = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_PROVAMT,AU_PROVCODE,"
    SQLQ = SQLQ & "AU_PAYROLL_ID," 'If glbSoroc Then
    SQLQ = SQLQ & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
    
    If Not glbSQL Then SQLQ = SQLQ & " IN '" & glbIHRAUDIT & "' [;PWD=petman;DATABASE=" & glbIHRAUDIT & "] "
    
    SQLQ = SQLQ & " SELECT 'M','N',ED_EMPNBR, " & txtProvNDollar
    SQLQ = SQLQ & ",ED_PROVCODE"
    SQLQ = SQLQ & ",ED_PAYROLL_ID" 'If glbSoroc Then
    SQLQ = SQLQ & ",ED_DIV,ED_PT, "
    SQLQ = SQLQ & Date_SQL(Now) & " As AU_LDATE, '"
    SQLQ = SQLQ & Time$ & "' As AU_LTIME, 'N', '" & glbUserID & "' As AU_LUSER FROM HREMP "
    
    gdbAdoIhr001.Execute SQLQ & SQLQW
       
    SQLQ = "UPDATE HREMP SET ED_PROVAMT = " & txtProvNDollar
    If Len(txtProvNCode.Text) > 0 Then
        SQLQ = SQLQ & ",ED_PROVCODE = '" & txtProvNCode & "' "
    End If
    gdbAdoIhr001.Execute SQLQ & SQLQW
        
    Screen.MousePointer = DEFAULT
    
    flgProvUpd = True
End If

Continue:

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(0).FloodShowPct = False

If flgTD1Upd Or flgProvUpd Then
    MsgBox IIf(flgTD1Upd And flgProvUpd, "TD1 & Prov. Dollar/Code", IIf(flgTD1Upd, "TD1 Dollar/Code", IIf(flgProvUpd, "Prov. Dollar/Code", ""))) & " Update completed."
End If

End Sub

Private Sub chkBasicProvAmt_Click(Value As Integer)
    If chkBasicProvAmt Then
        lblProvODollar.Enabled = False
        txtProvODollar.Enabled = False
        lblProvNDollar.Enabled = False
        
        'Get Basic Prov. Amount from Company Master
        txtProvNDollar.Text = get_BasicTaxAmounts("PC_PROVTAX")
        
        txtProvNDollar.Locked = True
    Else
        lblProvODollar.Enabled = True
        txtProvODollar.Enabled = True
        lblProvNDollar.Enabled = True
        txtProvNDollar.Locked = False
    End If
End Sub

Private Sub chkBasicTD1Amt_Click(Value As Integer)
    If chkBasicTD1Amt Then
        lblODollar.Enabled = False
        txtODollar.Enabled = False
        lblNDollar.Enabled = False
        
        'Get Basic Prov. Amount from Company Master
        txtNDollar.Text = get_BasicTaxAmounts("PC_FEDTAX")
        
        txtNDollar.Locked = True
    Else
        lblODollar.Enabled = True
        txtODollar.Enabled = True
        lblNDollar.Enabled = True
        txtNDollar.Locked = False
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUTD1DOLLAR"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUTD1DOLLAR"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call setRptCaption(Me)
Call INI_Controls(Me)

End Sub

Private Sub Form_LostFocus()
    MDIMain.panHelp(0).Caption = " "
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUTd1Dollar = Nothing
End Sub

Private Sub txtNCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtNDollar_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtOCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtODollar_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

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
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Private Sub Transfer_TaxAmount(xEmpnbr, oldValue, AmountField As Control)
    Dim HRChanges As New Collection
    Call isChanged_Field(HRChanges, oldValue, AmountField, True)
    Call Passing_Changes(HRChanges, Banking, "M", Date, xEmpnbr)
End Sub

