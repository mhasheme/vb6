VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmBENGRCopy 
   Caption         =   "Copy Benefit Group"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   15300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraWFCUptSignAppr 
      Height          =   3015
      Left            =   10920
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdStar3 
         Appearance      =   0  'Flat
         Caption         =   "&Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Tag             =   "Save changes made"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton cmdClos3 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
      Begin MSMask.MaskEdBox MskFiscalYear 
         DataField       =   "FiscalYear"
         DataSource      =   "data1"
         Height          =   315
         Left            =   2040
         TabIndex        =   25
         Tag             =   "01-High Dollars"
         Top             =   360
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
         Height          =   285
         Index           =   4
         Left            =   1725
         TabIndex        =   26
         Top             =   750
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "MarketLine"
         Height          =   285
         Index           =   5
         Left            =   5445
         TabIndex        =   27
         Tag             =   "00-Market Line - Code"
         Top             =   1110
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFML"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "BAND"
         Height          =   285
         Index           =   6
         Left            =   5445
         TabIndex        =   28
         Tag             =   "00-Band - Code"
         Top             =   1470
         Visible         =   0   'False
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "WFBD"
      End
      Begin MSMask.MaskEdBox MskApprLimit 
         DataField       =   "APPR_LIMIT"
         DataSource      =   "data1"
         Height          =   315
         Left            =   5760
         TabIndex        =   29
         Tag             =   "01-High Dollars"
         Top             =   1830
         Visible         =   0   'False
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   43
         Top             =   1830
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   3840
         TabIndex        =   42
         Top             =   1470
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblMarketLine 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Market Line"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3840
         TabIndex        =   41
         Top             =   1110
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Plant"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   40
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblFiscalYear 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame fraWFCPCopyByDiv 
      Height          =   2415
      Left            =   6720
      TabIndex        =   34
      Top             =   2400
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdClos2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   24
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdStar2 
         Appearance      =   0  'Flat
         Caption         =   "&Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Tag             =   "Save changes made"
         Top             =   1920
         Width           =   1095
      End
      Begin INFOHR_Controls.CodeLookup clpDivF 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Tag             =   "00-Division"
         Top             =   480
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpDivT 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Tag             =   "00-Division"
         Top             =   840
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   2280
         TabIndex        =   22
         Tag             =   "00-Enter Region Code"
         Top             =   1200
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Plant Business Unit"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   37
         Top             =   1200
         Width           =   1860
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   795
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "From Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.Frame fraWFCPosCopy 
      Height          =   2415
      Left            =   6720
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton cmdStart 
         Appearance      =   0  'Flat
         Caption         =   "&Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Tag             =   "Save changes made"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   1920
         Width           =   1095
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   8
         Tag             =   "00-Enter Region Code"
         Top             =   960
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         Height          =   285
         Left            =   2400
         TabIndex        =   9
         Tag             =   "00-Division"
         Top             =   1320
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Plant Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Plant Business Unit"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   960
         Width           =   1740
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To Plant"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   600
      End
      Begin VB.Label lbltitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "From Plant"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "BM_COMMENTS"
      Height          =   1305
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "00-Comments - free form"
      Top             =   2040
      Width           =   8565
   End
   Begin VB.TextBox txtPolicy 
      Appearance      =   0  'Flat
      DataField       =   "BM_POLICY"
      Height          =   315
      Left            =   3080
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "00-Policy Number"
      Top             =   1320
      Width           =   4215
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   12
      Top             =   4995
      Width           =   15300
      _Version        =   65536
      _ExtentX        =   26987
      _ExtentY        =   979
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
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Tag             =   "Save changes made"
         Top             =   30
         Width           =   1095
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   3
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin INFOHR_Controls.CodeLookup clpBGroupOld 
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Tag             =   "01-Benefit - Group Code"
      Top             =   360
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BGMF"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpBGroupNew 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Tag             =   "01-Benefit - Group Code"
      Top             =   840
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BGMF"
      MaxLength       =   10
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   32
      Left            =   360
      TabIndex        =   16
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      Caption         =   "Policy Number"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   15
      Top             =   1380
      Width           =   1020
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      Caption         =   "Existing Benefit Group Code"
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   14
      Top             =   405
      Width           =   1980
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      Caption         =   "New Benefit Group Code"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   13
      Top             =   885
      Width           =   1770
   End
End
Attribute VB_Name = "frmBENGRCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLQ As String

Private Sub clpCode_LostFocus(Index As Integer)
    If glbWFC Then  'Ticket #29846 Franks 03/07/2017
        'If Index = 4 Or Index = 5 Or Index = 6 Then 'Band or MarketLine
        '    If Len(clpCode(4).Text) > 0 And Len(clpCode(5).Text) > 0 And Len(clpCode(6).Text) > 0 Then
        '        MskApprLimit.Text = WFCSigningApprovalLimitGet(clpCode(4).Text, clpCode(6).Text, clpCode(5).Text, MskFiscalYear.Text)  'xPlant, xBand, xMarketline, Optional xYear
        '    End If
        'End If
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClos2_Click()
Unload Me
End Sub

Private Sub cmdClos3_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim X, sUSERID, sCompPass, ICompPass
Dim sql As String
Dim OPwd As String
Dim rsBASIC As New ADODB.Recordset
Dim rsBASICTemp As New ADODB.Recordset
Dim Msg As String, a%

    If Not CriCheck() Then
        Exit Sub
    End If
    
    Msg = "This program will Copy all the Benefit Group Master and Benefit Group Matrix records "
    Msg = Msg & "from Benefit Group code '" & clpBGroupOld.Text & "' into Benefit Group code '" & clpBGroupNew.Text & "'." & Chr(10) & Chr(10)
    Msg = Msg & "Are you sure you want to copy it?"
    
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then
        Exit Sub
    End If
            
    Call CopySecurity
                    
    Unload Me
    
End Sub

Private Function CriCheck()
Dim rsBeGrp As New ADODB.Recordset
Dim X%

CriCheck = False

If Len(clpBGroupOld.Text) = 0 Then
    MsgBox "Existing Benefit Group Code!", vbCritical, ""
    clpBGroupOld.SetFocus
    Exit Function
Else
    If clpBGroupOld.Caption = "Unassigned" Then
        MsgBox "Benefit Group Code must be valid"
        clpBGroupOld.SetFocus
        Exit Function
    End If
End If

If Len(clpBGroupNew.Text) = 0 Then
    MsgBox "New Benefit Group Code!", vbCritical, ""
    clpBGroupNew.SetFocus
    Exit Function
Else
    If clpBGroupNew.Caption = "Unassigned" Then
        MsgBox "Benefit Group Code must be valid"
        clpBGroupNew.SetFocus
        Exit Function
    End If
End If

'check the new code in Benefit Group Master
SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & clpBGroupNew.Text & "' "
rsBeGrp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsBeGrp.EOF Then
    MsgBox "Benefit Group Code '" & clpBGroupNew.Text & "' exists in Benefit Group Master" & Chr(10) & "Can not copy to this code"
    clpBGroupNew.SetFocus
    Exit Function
End If
rsBeGrp.Close


SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE BM_BENEFIT_GROUP = '" & clpBGroupNew.Text & "' "
rsBeGrp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsBeGrp.EOF Then
    MsgBox "Benefit Group Code '" & clpBGroupNew.Text & "' exists in Benefit Group Matrix" & Chr(10) & "Only copy this code into Benefit Group Master."
    'clpBGroupNew.SetFocus
    'Exit Function
End If
rsBeGrp.Close

CriCheck = True
End Function

Private Sub cmdStar2_Click()
Dim X, sUSERID, sCompPass, ICompPass
Dim sql As String
Dim OPwd As String
Dim rsBASIC As New ADODB.Recordset
Dim rsBASICTemp As New ADODB.Recordset
Dim Msg As String, a%

    If Len(clpDivF.Text) = 0 Then
        MsgBox "From Division is required"
        clpDivF.SetFocus
        Exit Sub
    Else
        If clpDivF.Caption = "Unassigned" Then
            MsgBox "Invalid Division Code"
            clpDivF.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpDivT.Text) = 0 Then
        MsgBox "To Division is required"
        clpDivT.SetFocus
        Exit Sub
    Else
        If clpDivT.Caption = "Unassigned" Then
            MsgBox "Invalid Division Code"
            clpDivT.SetFocus
            Exit Sub
        End If
    End If
    If clpDivF.Text = clpDivT.Text Then
            MsgBox "From Division cannot be same as To Division"
            clpDivT.SetFocus
            Exit Sub
    End If
    If Len(clpCode(3).Text) = 0 Then
        MsgBox "New Business Unit is required"
        clpCode(3).SetFocus
        Exit Sub
    Else
        If clpCode(3).Caption = "Unassigned" Then
            MsgBox "Invalid Business Unit Code"
            clpCode(3).SetFocus
            Exit Sub
        End If
    End If

    Msg = "This program will Copy all the Positions and Budgeted Positions " & Chr(10)
    Msg = Msg & "from '" & clpDivF.Text & "' to '" & clpDivT.Text & "'" & Chr(10) & Chr(10)
    Msg = Msg & "Are you sure you want to do it?"
    
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then
        Exit Sub
    End If
    
    cmdStart.Enabled = False
    cmdClose.Enabled = False
    'Call CopyPosAndBudget
    Call CopyPosAndBudByDiv
    
    'MsgBox " Finished! "
    
    cmdStart.Enabled = True
    cmdClose.Enabled = True
    
    Unload Me

End Sub

Private Sub cmdStar3_Click()
Dim X, sUSERID, sCompPass, ICompPass
Dim sql As String
Dim OPwd As String
Dim rsBASIC As New ADODB.Recordset
Dim rsBASICTemp As New ADODB.Recordset
Dim Msg As String, a%
Dim xTot As Long

    If Len(MskFiscalYear.Text) > 0 Then
        If Not IsNumeric(MskFiscalYear.Text) Then
            MsgBox "Invalid Year."
            MskFiscalYear.SetFocus
            Exit Sub
        Else
            If Not Len(MskFiscalYear.Text) = 4 Then
                MsgBox "Invalid Year."
                MskFiscalYear.SetFocus
                Exit Sub
            End If
        End If
    Else
        MsgBox "Year is a required field"
        MskFiscalYear.SetFocus
        Exit Sub
    End If

    If Len(clpCode(4).Text) = 0 Then
        MsgBox "Plant is required"
        clpCode(4).SetFocus
        Exit Sub
    Else
        If clpCode(4).Caption = "Unassigned" Then
            MsgBox "Invalid Plant Code"
            clpCode(4).SetFocus
            Exit Sub
        End If
    End If
    

    ''If Len(clpCode(5).Text) = 0 Then
    ''    MsgBox "Market Line is required"
    ''    clpCode(5).SetFocus
    ''    Exit Sub
    ''Else
    ''    If clpCode(5).Caption = "Unassigned" Then
    ''        MsgBox "Invalid Market Line Code"
    ''        clpCode(5).SetFocus
    ''        Exit Sub
    ''    End If
    ''End If
    ''If Len(clpCode(6).Text) = 0 Then
    ''    MsgBox "Band is required"
    ''    clpCode(6).SetFocus
    ''    Exit Sub
    ''Else
    ''    If clpCode(6).Caption = "Unassigned" Then
    ''        MsgBox "Invalid Band Code"
    ''        clpCode(6).SetFocus
    ''        Exit Sub
    ''    End If
    ''End If
    ''
    ''If Len(MskApprLimit.Text) > 0 Then
    ''    If Not IsNumeric(MskApprLimit.Text) Then
    ''            MsgBox "Invalid Signing Approval Limit."
    ''            MskApprLimit.SetFocus
    ''        Exit Sub
    ''    End If
    ''Else
    ''    MsgBox "Signing Approval Limit is a required field"
    ''    MskApprLimit.SetFocus
    ''    Exit Sub
    ''End If
    
    Msg = "This program will update the Signing Approval Limit " & Chr(10)
    Msg = Msg & "for these Positions based on the Salary Grid" & Chr(10) & Chr(10)
    'Msg = Msg & "with the same Plant, Market Line and Band" & Chr(10) & Chr(10)
    Msg = Msg & "Are you sure you want to do it?"
    
    a% = MsgBox(Msg, 36, "Confirm Update")
    If a% <> 6 Then
        Exit Sub
    End If
    
    cmdStar3.Enabled = False
    cmdClos3.Enabled = False
    'Call CopyPosAndBudByDiv
    
    Screen.MousePointer = HOURGLASS
    'xTot = WFCSigningApprovalLimitUpt(clpCode(4).Text, clpCode(6).Text, clpCode(5).Text, MskApprLimit.Text) 'Plant, Band, Marketline, Limit
    xTot = WFCSigningApproLimitUpt(MskFiscalYear.Text, clpCode(4).Text)
    Screen.MousePointer = DEFAULT
    
    If xTot = 0 Then
        MsgBox "No matching Position Master record found."
    Else
        MsgBox xTot & " Position Master record(s) updated."
    End If

    cmdStar3.Enabled = True
    cmdClos3.Enabled = True
    
    Unload Me

End Sub

Private Sub cmdStart_Click()
Dim X, sUSERID, sCompPass, ICompPass
Dim sql As String
Dim OPwd As String
Dim rsBASIC As New ADODB.Recordset
Dim rsBASICTemp As New ADODB.Recordset
Dim Msg As String, a%

    If Len(clpCode(0).Text) = 0 Then
        MsgBox "From Plant is required"
        clpCode(0).SetFocus
        Exit Sub
    Else
        If clpCode(0).Caption = "Unassigned" Then
            MsgBox "Invalid Plant Code"
            clpCode(0).SetFocus
            Exit Sub
        End If
    End If
    If Len(clpCode(1).Text) = 0 Then
        MsgBox "To Plant is required"
        clpCode(1).SetFocus
        Exit Sub
    Else
        If clpCode(1).Caption = "Unassigned" Then
            MsgBox "Invalid Plant Code"
            clpCode(1).SetFocus
            Exit Sub
        End If
    End If
    If clpCode(0).Text = clpCode(1).Text Then
            MsgBox "From Plant cannot be same as To Plant"
            clpCode(1).SetFocus
            Exit Sub
    End If
    If Len(clpCode(2).Text) = 0 Then
        MsgBox "New Plant Business Unit is required"
        clpCode(2).SetFocus
        Exit Sub
    Else
        If clpCode(2).Caption = "Unassigned" Then
            MsgBox "Invalid Business Unit Code"
            clpCode(2).SetFocus
            Exit Sub
        End If
    End If
    
    If Len(clpDiv.Text) = 0 Then
        MsgBox "New Plant Division is required"
        clpDiv.SetFocus
        Exit Sub
    Else
        If clpDiv.Caption = "Unassigned" Then
            MsgBox "Invalid Division Code"
            clpDiv.SetFocus
            Exit Sub
        End If
    End If

    
    Msg = "This program will Copy all the Positions and Budgeted Positions " & Chr(10)
    Msg = Msg & "from '" & clpCode(0).Text & "' to '" & clpCode(1).Text & "'" & Chr(10) & Chr(10)
    Msg = Msg & "Are you sure you want to do it?"
    
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then
        Exit Sub
    End If
    
    cmdStart.Enabled = False
    cmdClose.Enabled = False
    Call CopyPosAndBudget
    
    'MsgBox " Finished! "
    
    cmdStart.Enabled = True
    cmdClose.Enabled = True
    
    Unload Me
End Sub

Private Sub Form_Load()
Me.Width = 9645
Me.Height = 5535

If glbWFC Then
    If glbWFC_IPPopFormName = "WFCPosMasterAndBudCopy" Then 'Ticket #29438 Franks 12/01/2016
        Call WFCPosCopyScreenSetup
    End If
    If glbWFC_IPPopFormName = "WFCPosMasterAndBudCopByDiv" Then 'Ticket #29846 Franks 03/06/2017
        Call WFCPosCopyByDivSetup
    End If
    If glbWFC_IPPopFormName = "WFCUptSigningApproval" Then 'Ticket #29846 Franks 03/07/2017
        Call WFCUptSigningApprSetup
    End If
End If
Call INI_Controls(Me)
End Sub

Private Sub CopySecurity()
Dim X, sUSERID
Dim rsACCESS As New ADODB.Recordset
Dim rsINSERT As New ADODB.Recordset
Dim rsCopySecurity As New ADODB.Recordset
    
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & clpBGroupOld.Text & "' "
    rsCopySecurity.Open SQLQ, gdbAdoIhr001, adOpenStatic

    sql = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & clpBGroupNew.Text & "' "
    rsACCESS.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic

    MDIMain.panHelp(0).Caption = "Please wait while system copies Benefit Group..."

'    ''*** if employee does not exist in access security then copy security starts
    Do While Not rsCopySecurity.EOF
         rsACCESS.AddNew
         rsACCESS("BM_BENEFIT_GROUP") = clpBGroupNew.Text
         rsACCESS("BM_BCODE") = rsCopySecurity("BM_BCODE")
         rsACCESS("BM_EDATE") = rsCopySecurity("BM_EDATE")
         rsACCESS("BM_COVER") = rsCopySecurity("BM_COVER")
         rsACCESS("BM_AMT") = rsCopySecurity("BM_AMT")
         rsACCESS("BM_PPAMT") = rsCopySecurity("BM_PPAMT")
         rsACCESS("BM_UNITCOST") = rsCopySecurity("BM_UNITCOST")
         rsACCESS("BM_PCE") = rsCopySecurity("BM_PCE")
         rsACCESS("BM_PCC") = rsCopySecurity("BM_PCC")
         rsACCESS("BM_ECOST") = rsCopySecurity("BM_ECOST")
         rsACCESS("BM_CCOST") = rsCopySecurity("BM_CCOST")
         rsACCESS("BM_TCOST") = rsCopySecurity("BM_TCOST")
         rsACCESS("BM_MAXDOL") = rsCopySecurity("BM_MAXDOL")
         rsACCESS("BM_PREMIUM") = rsCopySecurity("BM_PREMIUM")
         rsACCESS("BM_PER") = rsCopySecurity("BM_PER")
         rsACCESS("BM_MTHCCOST") = rsCopySecurity("BM_MTHCCOST")
         rsACCESS("BM_MTHECOST") = rsCopySecurity("BM_MTHECOST")
         rsACCESS("BM_TAXBEN") = rsCopySecurity("BM_TAXBEN")
         rsACCESS("BM_SALARYDEPENDANT") = rsCopySecurity("BM_SALARYDEPENDANT")
         rsACCESS("BM_MINIMUM") = rsCopySecurity("BM_MINIMUM")
         rsACCESS("BM_FACTOR") = rsCopySecurity("BM_FACTOR")
         rsACCESS("BM_ROUND") = rsCopySecurity("BM_ROUND")
         rsACCESS("BM_MAXIMUM") = rsCopySecurity("BM_MAXIMUM")
         rsACCESS("BM_NEXTNEAREST") = rsCopySecurity("BM_NEXTNEAREST")
         rsACCESS("BM_TAXAMOUNT") = rsCopySecurity("BM_TAXAMOUNT")
         If Len(txtPolicy.Text) > 0 Then
            rsACCESS("BM_POLICY") = txtPolicy.Text
         End If
         rsACCESS("BM_WAITPERIOD") = rsCopySecurity("BM_WAITPERIOD")
         If Len(memComments.Text) > 0 Then
            rsACCESS("BM_COMMENTS") = memComments.Text
         End If
         rsACCESS("BM_PTAX") = rsCopySecurity("BM_PTAX")
         rsACCESS("BM_SORTCODE") = rsCopySecurity("BM_SORTCODE")
         rsACCESS("BM_EMPLOYEEID") = rsCopySecurity("BM_EMPLOYEEID")
         rsACCESS("BM_DWM") = rsCopySecurity("BM_DWM")
         rsACCESS("BM_CYTD") = rsCopySecurity("BM_CYTD")
         rsACCESS("BM_EYTD") = rsCopySecurity("BM_EYTD")
         rsACCESS("BM_PERORDOLL") = rsCopySecurity("BM_PERORDOLL")
         rsACCESS("BM_LDATE") = Date
         rsACCESS("BM_LTIME") = Time$
         rsACCESS("BM_LUSER") = glbUserID
         rsACCESS.Update
         rsCopySecurity.MoveNext
    Loop
    rsACCESS.Close
    rsCopySecurity.Close

    'Copying Benefit Group Matrix '
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE BM_BENEFIT_GROUP = '" & clpBGroupOld.Text & "' "
    rsCopySecurity.Open SQLQ, gdbAdoIhr001, adOpenStatic

    sql = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE BM_BENEFIT_GROUP = '" & clpBGroupNew.Text & "' "
    rsACCESS.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    
    If Not rsACCESS.EOF Then
        'found this code in the Matrix table, skip copy
        GoTo end_line
    End If
    Do While Not rsCopySecurity.EOF
         rsACCESS.AddNew
         rsACCESS("BM_BENEFIT_GROUP") = clpBGroupNew.Text
         rsACCESS("BM_DIV") = rsCopySecurity("BM_DIV")
         rsACCESS("BM_BENEFIT_ACCOUNT") = rsCopySecurity("BM_BENEFIT_ACCOUNT")
         rsACCESS("BM_BENEFIT_CLASS") = rsCopySecurity("BM_BENEFIT_CLASS")
         rsACCESS("BM_CERTIFICATE_PREFIX") = rsCopySecurity("BM_CERTIFICATE_PREFIX")
         rsACCESS("BM_USER_FIELD1") = rsCopySecurity("BM_USER_FIELD1")
         rsACCESS("BM_USER_FIELD2") = rsCopySecurity("BM_USER_FIELD2")
         rsACCESS("BM_EMP") = rsCopySecurity("BM_EMP")
         rsACCESS("BM_TERM_REASON") = rsCopySecurity("BM_TERM_REASON")
         rsACCESS("BM_WHICH_DATE") = rsCopySecurity("BM_WHICH_DATE")
         rsACCESS("BM_FROM_DATE") = rsCopySecurity("BM_FROM_DATE")
         rsACCESS("BM_TO_DATE") = rsCopySecurity("BM_TO_DATE")
         rsACCESS("BM_COMMENTS") = rsCopySecurity("BM_COMMENTS")
         rsACCESS("BM_LDATE") = Date
         rsACCESS("BM_LTIME") = Time$
         rsACCESS("BM_LUSER") = glbUserID
         rsACCESS.Update
         rsCopySecurity.MoveNext
    Loop
    rsACCESS.Close
    rsCopySecurity.Close
    
end_line:
    Screen.MousePointer = DEFAULT
    MDIMain.panHelp(0).Caption = "Benefit Group Code Copying Done"
    MsgBox "Benefit Group Copy completed successfully."

    Unload Me
End Sub

Private Sub CopyPosAndBudByDiv()
Dim X, sUSERID
Dim rsPos As New ADODB.Recordset
Dim rsAddPos As New ADODB.Recordset
Dim rsFrom As New ADODB.Recordset
Dim rsTo As New ADODB.Recordset
Dim I As Integer
Dim xTot As Integer
Dim xCurPos, xNewPos, xJob, xJobGRP, xNewPlant, xNewDiv, xRegion
Dim xMsg As String
Dim AppPath, xlsFileMat, buf

    'SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & clpCode(0).Text & "' "
    SQLQ = "SELECT * FROM HRJOB WHERE JB_DIV = '" & clpDivF.Text & "' "
    SQLQ = SQLQ & "AND NOT JB_STATUS = 'INAC' "
    SQLQ = SQLQ & "AND NOT LEFT(JB_DESCR,2) = 'Z ' "
    
    rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsPos.EOF Then
        MsgBox "Can not find any record."
        Exit Sub
    End If
    I = 0
    xTot = rsPos.RecordCount
    
    'MDIMain.panHelp(0).Caption = "Please wait while system copies Positions..."
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    
    AppPath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")
    xlsFileMat = AppPath & "WFCPositionCopyList.csv"
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    Open xlsFileMat For Output As #1
    'Print header line
    buf = """Plant"",""Division"",""Band"",""Position"",""Position Description """
    Print #1, buf
    
    xNewDiv = clpDivT.Text
    xNewPlant = getSectionByDiv(Left(clpDivT.Text, 4))
    xRegion = clpCode(3).Text
    Do While Not rsPos.EOF
        MDIMain.panHelp(0).FloodPercent = Int((I / xTot) * 100)
        I = I + 1
        DoEvents
        xCurPos = rsPos("JB_CODE")
        xJob = rsPos("JB_JOBCODE")
        xJobGRP = rsPos("JB_GRPCD")
        
        'New Position Code - begin
        xNewPos = getNewPosCode(xNewDiv, xJobGRP)
        'check if there is duplicate record based on the keys: Div + JOB + Pos Group
        SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & xNewPlant & "' " 'New Plant
        SQLQ = SQLQ & "AND JB_DIV = '" & xNewDiv & "' "
        SQLQ = SQLQ & "AND JB_JOBCODE = '" & xJob & "' "
        SQLQ = SQLQ & "AND JB_GRPCD = '" & xJobGRP & "' "
        SQLQ = SQLQ & "AND JB_CODE = '" & xNewPos & "' "
        If rsAddPos.State <> 0 Then rsAddPos.Close
        rsAddPos.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        If rsAddPos.EOF Then
            'xNewPos = getNewPosCode(xNewDiv, xJobGRP)
            'rsAddPos.AddNew
            
            rsAddPos.AddNew
            rsAddPos("JB_CODE") = Left(xNewPos, 25)
            'rsPos("")
            rsAddPos("JB_DESCR") = rsPos("JB_DESCR")
            rsAddPos("JB_DESCR2") = rsPos("JB_DESCR2")
            rsAddPos("JB_DIV") = Left(xNewDiv, 4)
            If Len(xRegion) > 0 Then rsAddPos("JB_REGION") = Left(xRegion, 20)
            If Len(xNewPlant) > 0 Then rsAddPos("JB_SECTION") = Left(xNewPlant, 4)
            rsAddPos("JB_MERCER_NO") = rsPos("JB_MERCER_NO")
            rsAddPos("JB_JOBCODE") = rsPos("JB_JOBCODE")
            rsAddPos("JB_POINTS") = rsPos("JB_POINTS")
            rsAddPos("JB_STATUS") = rsPos("JB_STATUS")
            rsAddPos("JB_GRPCD") = rsPos("JB_GRPCD")
            rsAddPos("JB_LEVEL") = rsPos("JB_LEVEL")
            rsAddPos("JB_BAND") = rsPos("JB_BAND")
            rsAddPos("JB_ORG") = rsPos("JB_ORG")
            rsAddPos("JB_MARKETLINE") = rsPos("JB_MARKETLINE")
            rsAddPos("JB_FEDGRP") = rsPos("JB_FEDGRP")
            'rsAddPos("JB_REPTAU") = rsPos("JB_REPTAU") 'Ticket #29955 Franks 03/20/2017 - "   Do not copy the RA #1 field
            rsAddPos("JB_REPTAU2") = rsPos("JB_REPTAU2")
            rsAddPos("JB_REPTAU3") = rsPos("JB_REPTAU3")
            rsAddPos("JB_REPTAU4") = rsPos("JB_REPTAU4")
            rsAddPos("JB_FTENUM") = rsPos("JB_FTENUM")
            rsAddPos("JB_FTETOTNU") = rsPos("JB_FTETOTNU")
            rsAddPos("JB_FTEHRS") = rsPos("JB_FTEHRS")
            rsAddPos("JB_FTETOTHR") = rsPos("JB_FTETOTHR")
            rsAddPos("JB_DHRS") = rsPos("JB_DHRS")
            rsAddPos("JB_LDATE") = Date
            rsAddPos("JB_LTIME") = Time$
            rsAddPos("JB_LUSER") = "PosCopy"
            rsAddPos("JB_SDATE") = Date ' rsPos("JB_SDATE")
            rsAddPos("JB_EDATE") = rsPos("JB_EDATE")
            rsAddPos("JB_USERDEF1") = rsPos("JB_USERDEF1")
            rsAddPos("JB_USERDEF2") = rsPos("JB_USERDEF2")
            rsAddPos("JB_POSTYPE") = rsPos("JB_POSTYPE")
            'rsAddPos("JB_APPR_LIMIT") = rsPos("JB_APPR_LIMIT")
            'Ticket #29955 Franks 03/20/2017
            rsAddPos("JB_APPR_LIMIT") = WFCSigningApprovalLimitGet(xNewPlant, rsPos("JB_BAND"), rsPos("JB_MARKETLINE")) '(xPlant, xBand, xMarketline, Optional xYear)
            rsAddPos.Update '

            'buf = """Plant"",""Division"",""Band"",""Position"",""Position Description """
            buf = """" & xNewPlant & """"
            buf = buf & ",""" & xNewDiv & """"
            If IsNull(rsPos("JB_BAND")) Then buf = buf & "," Else buf = buf & ",""" & rsPos("JB_BAND") & """"
            'If IsNull(rsPos("JB_MARKETLINE")) Then buf = buf & "," Else buf = buf & ",""" & rsPos("JB_MARKETLINE") & """"
            buf = buf & ",""" & xNewPos & """"
            buf = buf & ",""" & rsPos("JB_DESCR") & """"
            Print #1, buf
            
            
            Call WFCNextPosNoSetup("Ongoing")
            
            'New Budgeted Position - begin
            SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xCurPos & "' "
            SQLQ = SQLQ & "AND NOT JG_CURRENT = 0 "
            If rsFrom.State <> 0 Then rsFrom.Close
            rsFrom.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
            If Not rsFrom.EOF Then 'found it in the From Position
                'check if it can be found in the To Position
                SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xNewPos & "' "
                SQLQ = SQLQ & "AND NOT JG_CURRENT = 0 "
                If rsTo.State <> 0 Then rsTo.Close
                rsTo.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                If rsTo.EOF Then
                    rsTo.AddNew
                    rsTo("JG_COMPNO") = "001"
                    rsTo("JG_CODE") = xNewPos  'glbPos
                    rsTo("JG_BUDPOSNBR") = rsFrom("JG_BUDPOSNBR")
                    rsTo("JG_DIV") = xNewDiv  '
                    'If Len(xDeptT) = 0 Then
                        rsTo("JG_DEPTNO") = rsFrom("JG_DEPTNO") 'Null
                    'Else
                    '    rsTo("JG_DEPTNO") = xDeptT '
                    'End If
                    rsTo("JG_GLNO") = rsFrom("JG_GLNO")
                    rsTo("JG_NBRFIL") = 0 'rsFrom("JG_NBRFIL")
                    rsTo("JG_BUDGNBR") = rsFrom("JG_BUDGNBR")
                    rsTo("JG_FTENUM") = rsFrom("JG_FTENUM")
                    rsTo("JG_FTENUMFILL") = 0 'rsFrom("JG_FTENUMFILL")
                    rsTo("JG_FTENUMVACN") = rsFrom("JG_FTENUM") ' rsFrom("JG_FTENUMVACN")
                    rsTo("JG_FTEHRS") = rsFrom("JG_FTEHRS")
                    rsTo("JG_FTETOTHR") = rsFrom("JG_FTETOTHR")
                    rsTo("JG_LDATE") = Date
                    rsTo("JG_LTIME") = Time$
                    rsTo("JG_LUSER") = glbUserID
                    rsTo("JG_SECTION") = xNewPlant 'rsFrom("JG_SECTION")
                    rsTo("JG_YEAR") = rsFrom("JG_YEAR")
                    rsTo("JG_FRDATE") = rsFrom("JG_FRDATE")
                    rsTo("JG_TODATE") = rsFrom("JG_TODATE")
                    rsTo("JG_EFDATE") = rsFrom("JG_EFDATE")
                    rsTo("JG_JREASON") = rsFrom("JG_JREASON")
                    rsTo("JG_VACANCY_POS") = rsFrom("JG_VACANCY_POS")
                    rsTo("JG_CURRENT") = 1 ' rsFrom("JG_CURRENT")
                    rsTo.Update
                End If
                
            End If
            
            'New Budgeted Position - end
        Else
            'duplicate found, do nothing
        End If
        'New Position Code - end
        
        rsPos.MoveNext
    Loop

end_line:
    
    Close #1
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    Screen.MousePointer = DEFAULT
    'MDIMain.panHelp(0).Caption = "Positions Code Copying Done"
    xMsg = "Positions and Budgeted Positions Copy completed successfully."
    xMsg = xMsg & Chr(10) & Chr(10) & "Do not forget to update 'Reports To 1'. "
    xMsg = xMsg & Chr(10) & Chr(10) & "Please open this file to check which Positions were copied. "
    xMsg = xMsg & Chr(10) & xlsFileMat
    MsgBox xMsg

End Sub

Private Sub CopyPosAndBudget()
Dim X, sUSERID
Dim rsPos As New ADODB.Recordset
Dim rsAddPos As New ADODB.Recordset
Dim rsFrom As New ADODB.Recordset
Dim rsTo As New ADODB.Recordset
Dim I As Integer
Dim xTot As Integer
Dim xCurPos, xNewPos, xJob, xJobGRP, xNewPlant, xDiv, xRegion
Dim xMsg As String
Dim AppPath, xlsFileMat, buf

    'SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & clpBGroupOld.Text & "' "
    'rsCopySecurity.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & clpCode(0).Text & "' "
    SQLQ = SQLQ & "AND NOT JB_STATUS = 'INAC' "
    SQLQ = SQLQ & "AND NOT LEFT(JB_DESCR,2) = 'Z ' "
    SQLQ = SQLQ & "ORDER BY JB_DESCR"
    
    rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsPos.EOF Then
        MsgBox "Can not find any record."
        Exit Sub
    End If
    I = 0
    xTot = rsPos.RecordCount
    
    'MDIMain.panHelp(0).Caption = "Please wait while system copies Positions..."
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1

    AppPath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")
    xlsFileMat = AppPath & "WFCPositionCopyList.csv"
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    Open xlsFileMat For Output As #1
    'Print header line
    buf = """Plant"",""Division"",""Band"",""Position"",""Position Description """
    Print #1, buf
    
    
    xNewPlant = clpCode(1).Text
    xDiv = clpDiv.Text
    xRegion = clpCode(2).Text
    Do While Not rsPos.EOF
        MDIMain.panHelp(0).FloodPercent = Int((I / xTot) * 100)
        I = I + 1
        DoEvents
        xCurPos = rsPos("JB_CODE")
        xJob = rsPos("JB_JOBCODE")
        xJobGRP = rsPos("JB_GRPCD")
        
        'New Position Code - begin
        xNewPos = getNewPosCode(xDiv, xJobGRP)
        'check if there is duplicate record based on the keys: Div + JOB + Pos Group
        SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & xNewPlant & "' " 'New Plant
        SQLQ = SQLQ & "AND JB_DIV = '" & xDiv & "' "
        SQLQ = SQLQ & "AND JB_JOBCODE = '" & xJob & "' "
        SQLQ = SQLQ & "AND JB_GRPCD = '" & xJobGRP & "' "
        SQLQ = SQLQ & "AND JB_CODE = '" & xNewPos & "' "
        If rsAddPos.State <> 0 Then rsAddPos.Close
        rsAddPos.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        If rsAddPos.EOF Then
            'xNewPos = getNewPosCode(xDiv, xJobGRP)
            'rsAddPos.AddNew
            
            rsAddPos.AddNew
            rsAddPos("JB_CODE") = Left(xNewPos, 25)
            'rsPos("")
            rsAddPos("JB_DESCR") = rsPos("JB_DESCR")
            rsAddPos("JB_DESCR2") = rsPos("JB_DESCR2")
            rsAddPos("JB_DIV") = Left(xDiv, 4)
            If Len(xRegion) > 0 Then rsAddPos("JB_REGION") = Left(xRegion, 20)
            If Len(xNewPlant) > 0 Then rsAddPos("JB_SECTION") = Left(xNewPlant, 4)
            rsAddPos("JB_MERCER_NO") = rsPos("JB_MERCER_NO")
            rsAddPos("JB_JOBCODE") = rsPos("JB_JOBCODE")
            rsAddPos("JB_POINTS") = rsPos("JB_POINTS")
            rsAddPos("JB_STATUS") = rsPos("JB_STATUS")
            rsAddPos("JB_GRPCD") = rsPos("JB_GRPCD")
            rsAddPos("JB_LEVEL") = rsPos("JB_LEVEL")
            rsAddPos("JB_BAND") = rsPos("JB_BAND")
            rsAddPos("JB_ORG") = rsPos("JB_ORG")
            rsAddPos("JB_MARKETLINE") = rsPos("JB_MARKETLINE")
            rsAddPos("JB_FEDGRP") = rsPos("JB_FEDGRP")
            'Ticket #29955 Franks 03/20/2017 - "   Do not copy the RA #1 field
            'rsAddPos("JB_REPTAU") = rsPos("JB_REPTAU")
            rsAddPos("JB_REPTAU2") = rsPos("JB_REPTAU2")
            rsAddPos("JB_REPTAU3") = rsPos("JB_REPTAU3")
            rsAddPos("JB_REPTAU4") = rsPos("JB_REPTAU4")
            rsAddPos("JB_FTENUM") = rsPos("JB_FTENUM")
            rsAddPos("JB_FTETOTNU") = rsPos("JB_FTETOTNU")
            rsAddPos("JB_FTEHRS") = rsPos("JB_FTEHRS")
            rsAddPos("JB_FTETOTHR") = rsPos("JB_FTETOTHR")
            rsAddPos("JB_DHRS") = rsPos("JB_DHRS")
            rsAddPos("JB_LDATE") = Date
            rsAddPos("JB_LTIME") = Time$
            rsAddPos("JB_LUSER") = "PosCopy"
            rsAddPos("JB_SDATE") = Date ' rsPos("JB_SDATE")
            rsAddPos("JB_EDATE") = rsPos("JB_EDATE")
            rsAddPos("JB_USERDEF1") = rsPos("JB_USERDEF1")
            rsAddPos("JB_USERDEF2") = rsPos("JB_USERDEF2")
            rsAddPos("JB_POSTYPE") = rsPos("JB_POSTYPE")
            'rsAddPos("JB_APPR_LIMIT") = rsPos("JB_APPR_LIMIT")
            'Ticket #29955 Franks 03/20/2017
            rsAddPos("JB_APPR_LIMIT") = WFCSigningApprovalLimitGet(xNewPlant, rsPos("JB_BAND"), rsPos("JB_MARKETLINE")) '(xPlant, xBand, xMarketline, Optional xYear)
            rsAddPos.Update '
            
            'buf = """Plant"",""Division"",""Band"",""Position"",""Position Description """
            buf = """" & xNewPlant & """"
            buf = buf & ",""" & xDiv & """"
            If IsNull(rsPos("JB_BAND")) Then buf = buf & "," Else buf = buf & ",""" & rsPos("JB_BAND") & """"
            'If IsNull(rsPos("JB_MARKETLINE")) Then buf = buf & "," Else buf = buf & ",""" & rsPos("JB_MARKETLINE") & """"
            buf = buf & ",""" & xNewPos & """"
            buf = buf & ",""" & rsPos("JB_DESCR") & """"
            Print #1, buf
            
            Call WFCNextPosNoSetup("Ongoing")
            
            'New Budgeted Position - begin
            SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xCurPos & "' "
            SQLQ = SQLQ & "AND NOT JG_CURRENT = 0 "
            If rsFrom.State <> 0 Then rsFrom.Close
            rsFrom.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
            If Not rsFrom.EOF Then 'found it in the From Position
                'check if it can be found in the To Position
                SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xNewPos & "' "
                SQLQ = SQLQ & "AND NOT JG_CURRENT = 0 "
                If rsTo.State <> 0 Then rsTo.Close
                rsTo.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                If rsTo.EOF Then
                    rsTo.AddNew
                    rsTo("JG_COMPNO") = "001"
                    rsTo("JG_CODE") = xNewPos  'glbPos
                    rsTo("JG_BUDPOSNBR") = rsFrom("JG_BUDPOSNBR")
                    rsTo("JG_DIV") = xDiv  '
                    'If Len(xDeptT) = 0 Then
                        rsTo("JG_DEPTNO") = rsFrom("JG_DEPTNO") 'Null
                    'Else
                    '    rsTo("JG_DEPTNO") = xDeptT '
                    'End If
                    rsTo("JG_GLNO") = rsFrom("JG_GLNO")
                    rsTo("JG_NBRFIL") = 0 'rsFrom("JG_NBRFIL")
                    rsTo("JG_BUDGNBR") = rsFrom("JG_BUDGNBR")
                    rsTo("JG_FTENUM") = rsFrom("JG_FTENUM")
                    rsTo("JG_FTENUMFILL") = 0 'rsFrom("JG_FTENUMFILL")
                    rsTo("JG_FTENUMVACN") = rsFrom("JG_FTENUM") ' rsFrom("JG_FTENUMVACN")
                    rsTo("JG_FTEHRS") = rsFrom("JG_FTEHRS")
                    rsTo("JG_FTETOTHR") = rsFrom("JG_FTETOTHR")
                    rsTo("JG_LDATE") = Date
                    rsTo("JG_LTIME") = Time$
                    rsTo("JG_LUSER") = glbUserID
                    rsTo("JG_SECTION") = xNewPlant 'rsFrom("JG_SECTION")
                    rsTo("JG_YEAR") = rsFrom("JG_YEAR")
                    rsTo("JG_FRDATE") = rsFrom("JG_FRDATE")
                    rsTo("JG_TODATE") = rsFrom("JG_TODATE")
                    rsTo("JG_EFDATE") = rsFrom("JG_EFDATE")
                    rsTo("JG_JREASON") = rsFrom("JG_JREASON")
                    rsTo("JG_VACANCY_POS") = rsFrom("JG_VACANCY_POS")
                    rsTo("JG_CURRENT") = 1 ' rsFrom("JG_CURRENT")
                    rsTo.Update
                End If
                
            End If
            
            'New Budgeted Position - end
        Else
            'duplicate found, do nothing
        End If
        'New Position Code - end
        
        rsPos.MoveNext
    Loop
    
    Close #1
    
end_line:
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    Screen.MousePointer = DEFAULT
    'MDIMain.panHelp(0).Caption = "Positions Code Copying Done"
    xMsg = "Positions and Budgeted Positions Copy completed successfully."
    xMsg = xMsg & Chr(10) & Chr(10) & "Do not forget to update 'Reports To 1'. "
    xMsg = xMsg & Chr(10) & Chr(10) & "Please open this file to check which Positions were copied. "
    xMsg = xMsg & Chr(10) & xlsFileMat
    
    MsgBox xMsg
End Sub

Private Sub WFCPosCopyScreenSetup()
    Me.Caption = "Copy Positions By Plant"
    panControls.Visible = False
    Me.Width = 5985 + 200
    Me.Height = 3525
    fraWFCPosCopy.Left = 0
    fraWFCPosCopy.Top = 0
    fraWFCPosCopy.Width = 7695
    fraWFCPosCopy.Height = 3255
    fraWFCPosCopy.BorderStyle = 0
    fraWFCPosCopy.Visible = True
End Sub

Private Sub WFCPosCopyByDivSetup()
    Me.Caption = "Copy Positions By Division"
    panControls.Visible = False
    Me.Width = 5985 + 200
    Me.Height = 3525
    fraWFCPCopyByDiv.Left = 0
    fraWFCPCopyByDiv.Top = 0
    fraWFCPCopyByDiv.Width = 7695
    fraWFCPCopyByDiv.Height = 3255
    fraWFCPCopyByDiv.BorderStyle = 0
    fraWFCPCopyByDiv.Visible = True
End Sub

Private Sub WFCUptSigningApprSetup()
    Me.Caption = "Update Signing Approval"
    panControls.Visible = False
    Me.Width = 5985 + 200
    Me.Height = 3525
    fraWFCUptSignAppr.Left = 0
    fraWFCUptSignAppr.Top = 0
    fraWFCUptSignAppr.Width = 7695
    fraWFCUptSignAppr.Height = 3255
    fraWFCUptSignAppr.BorderStyle = 0
    fraWFCUptSignAppr.Visible = True
End Sub
