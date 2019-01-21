VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCHANGETYPE 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Change Type"
   ClientHeight    =   6930
   ClientLeft      =   1125
   ClientTop       =   795
   ClientWidth     =   6900
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6930
   ScaleWidth      =   6900
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   17
      Left            =   2640
      TabIndex        =   14
      Tag             =   "Completed"
      Top             =   4920
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   16
      Left            =   2640
      TabIndex        =   13
      Tag             =   "Completed"
      Top             =   4590
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   14
      Left            =   2640
      TabIndex        =   16
      Tag             =   "Completed"
      Top             =   5595
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   15
      Left            =   2640
      TabIndex        =   15
      Tag             =   "Completed"
      Top             =   5250
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      Tag             =   "Completed"
      Top             =   1290
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Tag             =   "Completed"
      Top             =   630
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   13
      Left            =   6120
      TabIndex        =   31
      Tag             =   "Completed"
      Top             =   5925
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   12
      Left            =   2640
      TabIndex        =   12
      Tag             =   "Completed"
      Top             =   4260
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   11
      Left            =   2640
      TabIndex        =   11
      Tag             =   "Completed"
      Top             =   3930
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   10
      Left            =   2640
      TabIndex        =   10
      Tag             =   "Completed"
      Top             =   3600
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   9
      Left            =   2640
      TabIndex        =   9
      Tag             =   "Completed"
      Top             =   3270
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   8
      Left            =   2640
      TabIndex        =   8
      Tag             =   "Completed"
      Top             =   2940
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   7
      Left            =   2640
      TabIndex        =   7
      Tag             =   "Completed"
      Top             =   2610
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   6
      Left            =   2640
      TabIndex        =   6
      Tag             =   "Completed"
      Top             =   2280
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   5
      Left            =   2640
      TabIndex        =   5
      Tag             =   "Completed"
      Top             =   1950
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Tag             =   "Completed"
      Top             =   1620
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   2
      Tag             =   "Completed"
      Top             =   960
      Width           =   285
   End
   Begin VB.CheckBox chkCompleted 
      DataField       =   "EC_INJURED_ONLINE"
      Height          =   195
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Tag             =   "Completed"
      Top             =   300
      Width           =   285
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   600
      Left            =   0
      TabIndex        =   19
      Top             =   6330
      Width           =   6900
      _Version        =   65536
      _ExtentX        =   12171
      _ExtentY        =   1058
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
         Left            =   1500
         TabIndex        =   18
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
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
         Left            =   420
         TabIndex        =   17
         Tag             =   "Select this Charge Code "
         Top             =   120
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   2100
         Top             =   -900
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rept. Authority 1"
      Height          =   195
      Index           =   17
      Left            =   420
      TabIndex        =   38
      Top             =   4927
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   195
      Index           =   16
      Left            =   420
      TabIndex        =   37
      Top             =   4598
      Width           =   1755
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Smoker"
      Height          =   195
      Index           =   14
      Left            =   420
      TabIndex        =   36
      Top             =   5595
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      Height          =   195
      Index           =   15
      Left            =   420
      TabIndex        =   35
      Top             =   5256
      Width           =   960
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   34
      Top             =   1308
      Width           =   630
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   33
      Top             =   650
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      Height          =   195
      Index           =   13
      Left            =   3900
      TabIndex        =   32
      Top             =   5925
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   30
      Top             =   4269
      Width           =   450
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      Height          =   195
      Index           =   11
      Left            =   420
      TabIndex        =   29
      Top             =   3940
      Width           =   540
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      Height          =   195
      Index           =   10
      Left            =   420
      TabIndex        =   28
      Top             =   3611
      Width           =   510
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      Height          =   195
      Index           =   9
      Left            =   420
      TabIndex        =   27
      Top             =   3282
      Width           =   420
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      Height          =   195
      Index           =   8
      Left            =   420
      TabIndex        =   26
      Top             =   2953
      Width           =   615
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTE# Hours"
      Height          =   195
      Index           =   7
      Left            =   420
      TabIndex        =   25
      Top             =   2624
      Width           =   870
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTE#"
      Height          =   195
      Index           =   6
      Left            =   420
      TabIndex        =   24
      Top             =   2295
      Width           =   405
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   195
      Index           =   5
      Left            =   420
      TabIndex        =   23
      Top             =   1966
      Width           =   825
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   22
      Top             =   1637
      Width           =   555
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Group"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   21
      Top             =   979
      Width           =   975
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   20
      Top             =   300
      Width           =   165
   End
End
Attribute VB_Name = "frmCHANGETYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNoRecords%
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim I As Integer

Private Sub chkCompleted_Click(Index As Integer)
    If Index = 0 Then
        If glbWFC Then 'Ticket #21118 Franks 10/31/2011
            For I = 1 To 15     'Ticket #28794 - Opening up the Marital Status for everyone
                chkCompleted(I).Value = chkCompleted(0).Value
            Next
        Else
            'Ticket #22304 - Changed from 13 to 12 because no Salary
            For I = 1 To 12
                chkCompleted(I).Value = chkCompleted(0).Value
            Next
            'Ticket #28794 - Opening up the Marital Status for everyone
            For I = 15 To 15
                chkCompleted(I).Value = chkCompleted(0).Value
            Next
        End If
        'cmdSelect.Enabled = chkCompleted(0).Value
        
        For I = 16 To 17 'Ticket #27553 Franks 09/22/2015
            chkCompleted(I).Value = chkCompleted(0).Value
        Next
    Else
        
    End If
    cmdSelect.Enabled = True
End Sub

Private Sub cmdClose_Click()

'glbCode = ""


Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdSelect_Click()
Dim x
glbCode = ""
'Ticket #27553 Franks 09/22/2015
For I = 1 To 17 '15 '13 'Ticket #21118 Franks 10/31/2011
    'Ticket #22304 - No Salary
    If I <> 13 Then     'Ticket #22559 - It used to be 12 which is for Status. Salary is 13.
        If chkCompleted(I).Value = 1 Then
            If Len(glbCode) > 1 Then
                glbCode = glbCode & "," & lblTitle(I).Caption
            Else
                glbCode = lblTitle(I).Caption
            End If
        End If
    End If
Next
'If chkCompleted(13).Value = 1 Then
'    If Len(glbCode) > 1 Then
'        glbCode = glbCode & ","""""
'    Else
'        glbCode = """"""
'    End If
'End If

Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Dim xStr

End Sub

Private Sub Form_Load()
Dim SQLQ, intCount
glbOnTop = "FRMCHANGETYPE"
intCount = 0
Screen.MousePointer = DEFAULT
If glbWFC Then 'Ticket #21118 Franks 10/31/2011
    lblTitle(14).Visible = True
    lblTitle(15).Visible = True
    chkCompleted(14).Visible = True
    chkCompleted(15).Visible = True     'Ticket #28794 - Opening up the Marital Status for everyone
    For I = 0 To 15
        lblTitle(I).Caption = lStr(lblTitle(I).Caption)
        If InStr(1, glbCode, lblTitle(I).Caption) > 0 Then
            chkCompleted(I).Value = 1
            intCount = intCount + 1
        End If
    Next
    If intCount = 15 Then chkCompleted(0).Value = 1
    
Else
    'Ticket #22304 - Changed from 13 to 12 because no Salary
    For I = 0 To 12
        lblTitle(I).Caption = lStr(lblTitle(I).Caption)
        If InStr(1, glbCode, lblTitle(I).Caption) > 0 Then
            chkCompleted(I).Value = 1
            intCount = intCount + 1
        End If
    Next
    
    'Ticket #22304 - Changed from 13 to 12 because no Salary
    If intCount = 12 Then chkCompleted(0).Value = 1
    
    'Ticket #28794 - Opening up the Marital Status for everyone
    For I = 15 To 15
        lblTitle(I).Caption = lStr(lblTitle(I).Caption)
        If InStr(1, glbCode, lblTitle(I).Caption) > 0 Then
            chkCompleted(I).Value = 1
            intCount = intCount + 1
        End If
    Next
End If

For I = 16 To 17 'Ticket #27553 Franks 09/22/2015
    lblTitle(I).Caption = lStr(lblTitle(I).Caption)
    If InStr(1, glbCode, lblTitle(I).Caption) > 0 Then
        chkCompleted(I).Value = 1
        intCount = intCount + 1
    End If
Next

If intCount > 0 Then cmdSelect.Enabled = True

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub
