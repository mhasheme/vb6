VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHistory 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee History"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
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
   Icon            =   "fehistory.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7905
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNewValue 
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
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   3735
   End
   Begin INFOHR_Controls.CodeLookup clpNewValue 
      DataField       =   "NewValue"
      Height          =   285
      Left            =   1740
      TabIndex        =   6
      Tag             =   "01-Skills - Code"
      Top             =   4080
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSK"
      MaxLength       =   20
   End
   Begin VB.ComboBox ComMStatOld 
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
      ItemData        =   "fehistory.frx":030A
      Left            =   2160
      List            =   "fehistory.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Tag             =   "Marital Status"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox ComMStatNew 
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
      ItemData        =   "fehistory.frx":030E
      Left            =   2160
      List            =   "fehistory.frx":0310
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Tag             =   "Marital Status"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtOldValue 
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
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CheckBox chkChgSalary 
      Alignment       =   1  'Right Justify
      Caption         =   "Salary Change"
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
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EE_LUSER"
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
      Left            =   6450
      MaxLength       =   25
      TabIndex        =   26
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EE_LTIME"
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
      Left            =   4770
      MaxLength       =   25
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EE_LDATE"
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
      Index           =   0
      Left            =   3090
      MaxLength       =   25
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.ComboBox comPayPer 
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
      Left            =   4170
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "01-Choose annum or hour"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox comType 
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
      Left            =   2040
      TabIndex        =   2
      Text            =   "comType"
      Top             =   3360
      Width           =   2055
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehistory.frx":0312
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "fehistory.frx":0326
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
   Begin INFOHR_Controls.DateLookup dlpCHGDate 
      DataField       =   "EE_CHGDATE"
      Height          =   285
      Left            =   1725
      TabIndex        =   1
      Tag             =   "40-Enter date that skill was acquired"
      Top             =   3000
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpOldValue 
      DataField       =   "OldValue"
      Height          =   285
      Left            =   1740
      TabIndex        =   5
      Tag             =   "01-Skills - Code"
      Top             =   3720
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSK"
      MaxLength       =   20
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7440
      Top             =   7200
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   20
      Top             =   7245
      Width           =   9660
      _Version        =   65536
      _ExtentX        =   17039
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
         Left            =   6495
         Top             =   120
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
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9660
      _Version        =   65536
      _ExtentX        =   17039
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
         Left            =   6840
         TabIndex        =   30
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   160
         Width           =   1005
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
         TabIndex        =   14
         Top             =   135
         Width           =   1245
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
         TabIndex        =   13
         Top             =   135
         Width           =   720
      End
   End
   Begin MSMask.MaskEdBox medsalary 
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Tag             =   "21-Enter salary"
      Top             =   5055
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
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
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOldFTE 
      DataField       =   "OldValue"
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Tag             =   "10-Full - time equivalency"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
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
      Format          =   "##0.00"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medNewFTE 
      DataField       =   "NewValue"
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Tag             =   "10-FTE Hours worked per year"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
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
      Format          =   "###0.00"
      PromptChar      =   "_"
   End
   Begin VB.Label lblType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "HISTYPE"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2580
      TabIndex        =   29
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label lblSalCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SalCode"
      DataField       =   "EE_SALCD"
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5640
      TabIndex        =   28
      Top             =   5100
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EE_COMPNO"
      DataSource      =   "Data1"
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
      Left            =   540
      TabIndex        =   27
      Top             =   6480
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   23
      Top             =   5100
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Per"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   3780
      TabIndex        =   22
      Top             =   5100
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblNewValue 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "New Value"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   21
      Top             =   4110
      Width           =   930
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Date"
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
      Left            =   300
      TabIndex        =   19
      Top             =   3030
      Width           =   945
   End
   Begin VB.Label lbType 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Change Type"
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
      Left            =   300
      TabIndex        =   18
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label lblOldValue 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Value"
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
      Left            =   300
      TabIndex        =   17
      Top             =   3780
      Width           =   1155
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EE_EMPNBR"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   6420
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer
Dim fglbNewSalRec%
Dim glbNew
Dim fglbNew
Dim RSDATA As New ADODB.Recordset
Dim FRS As ADODB.Recordset
Dim Ctrl As Control
Dim fglbSQL

Private Function chkValues()
Dim Values As String, Valuesvl As String, ValuesDte As String
Dim SQLQ As String, Msg As String

chkValues = False

On Error GoTo chkValue_Err

If Not (Me.chkChgSalary.Value = 1 Or comType.Text = lStr("FTE# Hours") Or comType.Text = lStr("FTE#")) Then
    If Len(clpNewValue.Text) < 1 Then
        MsgBox "New value code is a required field"
        If clpNewValue.Visible Then clpNewValue.SetFocus
        Exit Function
    End If
    
    If clpNewValue.Visible Then 'Ticket #21118 Franks 10/28/2011
        If clpNewValue.Caption = "Unassigned" Then
            MsgBox "New value code must be valid"
            If clpNewValue.Visible Then clpNewValue.SetFocus
            Exit Function
        End If
    End If
End If

If Len(dlpCHGDate.Text) < 1 Then
    dlpCHGDate.Text = Format(Now, "short date")
End If

If Not IsDate(dlpCHGDate.Text) Then
    MsgBox "Change Date is not a valid date"
    dlpCHGDate.SetFocus
    Exit Function
End If
                                  

If glbWFC Then 'Ticket #21119 Franks 11/14/2011
    If Len(glbWFCNGSSubGroup) > 0 Then 'NGS only
        If comType = "Smoker" Then
            If fglbNew Then
                MsgBox "Can not add Smoker record for NGS employees."
                Exit Function
            Else
                MsgBox "Can not edit Smoker record for NGS employees."
                Exit Function 'Can not edit and delete the record
            End If
        End If
    End If
End If

chkValues = True

Exit Function

chkValue_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkValue", "HREMPHIS", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err
'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove ADO
If Not (RSDATA.EOF And RSDATA.BOF) Then
    RSDATA.CancelUpdate
End If
Call Display_Value

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMPHIS", "Cancel")
Call RollBack

End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "frmEHistory" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Values As String, Valuesvl As String, ValuesDte As String
Dim SQLQ As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

If glbWFC Then 'Ticket #21118 Franks 10/28/2011
    If Len(glbWFCNGSSubGroup) > 0 Then 'NGS only
        If comType = "Smoker" Then
            MsgBox "Can not delete Smoker record for NGS employees."
            Exit Sub 'Can not edit and delete the record
        End If
    End If
End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001.CommitTrans
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREMPHIS", "Delete")
Call RollBack '23July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err

Call Set_Control("B", Me)

lbType.Visible = True
lblType.Visible = True
lblNewValue.Visible = True
lblOldValue.Visible = True
clpOldValue.Enabled = True
clpNewValue.Enabled = True
comType.Enabled = True
dlpCHGDate.Enabled = True
comType.Text = ""
chkChgSalary.Enabled = False
chkChgSalary.Value = 0
lblTitle(5).Visible = False
lblTitle(6).Visible = False
medsalary.Visible = False
comPayPer.Visible = False
'rsDATA.AddNew

clpOldValue.Text = ""
clpNewValue.Text = ""
txtOldValue.Text = ""
txtNewValue.Text = ""


dlpCHGDate.Text = Format(Now, "short date")  ', short date)
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

'Me.clpNewValue.Enabled = True
'Me.clpNewValue.SetFocus

glbNew = True
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREMPHIS", "Add")
Resume Next
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x
On Error GoTo Add_Err


If Not chkValues() Then Exit Sub


Call UpdUStats(Me)

Dim xID
If fglbNew Then
    xID = 0
Else
    xID = Data1.Recordset("EE_ID")
End If

Dim rsSave As New ADODB.Recordset
Dim xCtrl As Control
Dim SQLQ
'Call Set_Control("U", Me, rsDATA)


Dim xFieldPartName
Select Case comType
Case lStr("Administered By")
    xFieldPartName = "ADMINBY"
Case lStr("Benefit Group")
    xFieldPartName = "BENEGROUP"
Case lStr("Department")
    xFieldPartName = "DEPT"
Case lStr("Division")
    xFieldPartName = "DIV"
Case lStr("Category")
    xFieldPartName = "PT"
Case lStr("FTE#")
    xFieldPartName = "FTE"
Case lStr("FTE# Hours")
    xFieldPartName = "FTEHR"
Case lStr("Location")
    xFieldPartName = "LOC"
Case lStr("Union")
    xFieldPartName = "ORG"
Case lStr("Region")
    xFieldPartName = "REGION"
'Case "Salary"
'    xFieldPartName = "SALARY"
Case lStr("Section")
    xFieldPartName = "SECTION"
Case lStr("Status")
    xFieldPartName = "STAT"
Case ("Smoker") 'Ticket #21118 Franks 10/28/2011
    xFieldPartName = "SMOKER"
Case ("Marital Status") 'Ticket #21118 Franks 10/28/2011
    xFieldPartName = "MSTAT"
End Select


If glbtermopen Then
    SQLQ = "SELECT * FROM Term_HREMPHIS WHERE EE_ID=" & xID
    rsSave.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockOptimistic
    If fglbNew Then rsSave.AddNew
    
    rsSave("EE_COMPNO") = lblCNum
    rsSave("EE_EMPNBR") = lblEEID
    rsSave("EE_LDATE") = Updstats(0)
    rsSave("EE_LTIME") = Updstats(1)
    rsSave("EE_LUSER") = Updstats(2)
    
    rsSave("EE_CHGDATE") = dlpCHGDate
    If chkChgSalary.Value = 0 Then  'Hemu - Begin - since the SALARY fieldpartname is not assigned above for Salary
        rsSave("EE_NEW" & xFieldPartName) = clpNewValue
        If clpOldValue <> "" Then
            rsSave("EE_OLD" & xFieldPartName) = clpOldValue
        Else
            rsSave("EE_OLD" & xFieldPartName) = Null
        End If
    End If  'Hemu - End
    If chkChgSalary Then
        rsSave("EE_SALARY") = Val(medsalary)
        rsSave("EE_SALCD") = Left(comPayPer, 1)
    Else
        rsSave("EE_SALARY") = 0
        rsSave("EE_SALCD") = " "
    End If
    
    rsSave!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsSave.Update
    gdbAdoIhr001X.CommitTrans
Else
    SQLQ = "SELECT * FROM HREMPHIS WHERE EE_ID=" & xID
    rsSave.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If fglbNew Then rsSave.AddNew
    
    rsSave("EE_COMPNO") = lblCNum
    rsSave("EE_EMPNBR") = lblEEID
    rsSave("EE_LDATE") = Updstats(0)
    rsSave("EE_LTIME") = Updstats(1)
    rsSave("EE_LUSER") = Updstats(2)
    
    rsSave("EE_CHGDATE") = dlpCHGDate
    If chkChgSalary.Value = 0 Then  'Hemu - Begin - since the SALARY fieldpartname is not assigned above for Salary
        rsSave("EE_NEW" & xFieldPartName) = clpNewValue
        If clpOldValue <> "" Then
            rsSave("EE_OLD" & xFieldPartName) = clpOldValue
        Else
            rsSave("EE_OLD" & xFieldPartName) = Null
        End If
    End If  'Hemu - End
    If chkChgSalary Then
        rsSave("EE_SALARY") = Val(medsalary)
        rsSave("EE_SALCD") = Left(comPayPer, 1)
    Else
        rsSave("EE_SALARY") = 0
        rsSave("EE_SALCD") = " "
    End If
    
    gdbAdoIhr001.BeginTrans
    rsSave.Update
    gdbAdoIhr001.CommitTrans
End If
rsSave.Close

Data1.Refresh
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
If NextFormIF("History") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 424 Or Err = 438 Then
    Resume Next
ElseIf Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    Data1.Recordset.CancelUpdate
    Data1.Recordset.Resync
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMPHIS", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Employee History"
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

RHeading = lblEEName & "'s Employee History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub chkChgSalary_Click()
    'Ticket #18435 Frank 04/29/2010
    'Samuel doesn't want to see salary on this screen. Jerry asked to hide all salary fields
    'lblTitle(5).Visible = chkChgSalary <> 0
    'lblTitle(6).Visible = chkChgSalary <> 0
    'medsalary.Visible = chkChgSalary <> 0
    'comPayPer.Visible = chkChgSalary <> 0
    '
    'lbType.Visible = chkChgSalary = 0
    'lblType.Visible = chkChgSalary = 0
    'lblNewValue.Visible = chkChgSalary = 0
    'lblOldValue.Visible = chkChgSalary = 0
    'clpOldValue.Visible = chkChgSalary = 0
    'clpNewValue.Visible = chkChgSalary = 0
    'txtNewValue.Visible = chkChgSalary = 0
    'txtOldValue.Visible = chkChgSalary = 0
    '
    'comType.Visible = chkChgSalary = 0
End Sub

''Private Sub clpOldValue_Change()
''If glbWFC Then 'Ticket #21118 Franks 10/28/2011
''    If Not Data1.Recordset.EOF Then
''        If Data1.Recordset("HISTYPE") = "Marital Status" Then
''            ComMStatOld.ListIndex = GetComMStatInx(clpOldValue.Text)
''        End If
''    End If
''End If
''End Sub
''Private Sub clpNewValue_Change()
''If glbWFC Then 'Ticket #21118 Franks 10/28/2011
''    If Not Data1.Recordset.EOF Then
''        If Data1.Recordset("HISTYPE") = "Marital Status" Then
''            ComMStatNew.ListIndex = GetComMStatInx(clpNewValue.Text)
''        End If
''    End If
''End If
''End Sub
''Private Sub ComMStatOld_Click() 'Ticket #21118 Franks 10/28/2011
''    If ComMStatOld = "Partner" Or ComMStatOld = "Same-Sex" Then
''        clpOldValue.Text = UCase(Right(ComMStatOld.Text, 1))
''        txtOldValue.Text = UCase(Right(ComMStatOld.Text, 1))
''    ElseIf ComMStatOld = "Separated" Then
''        clpOldValue.Text = UCase(Mid(ComMStatOld.Text, 4, 1))
''        txtOldValue.Text = UCase(Right(ComMStatOld.Text, 1))
''    Else
''        clpOldValue.Text = Left(ComMStatOld.Text, 1)
''        txtOldValue.Text = UCase(Right(ComMStatOld.Text, 1))
''    End If
''End Sub

Private Sub comType_Click()
'clpNewValue.Visible = True
'clpOldValue.Visible = True
Select Case comType
Case lStr("Administered By")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDAB"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDAB"
Case lStr("Benefit Group")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "BGMF"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "BGMF"
Case lStr("Department")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = Department
    clpOldValue.TablName = "n/a"
    clpNewValue.LookupType = Department
    clpNewValue.TablName = "n/a"
Case lStr("Division")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = Division
    clpOldValue.TablName = "n/a"
    clpNewValue.LookupType = Division
    clpNewValue.TablName = "n/a"
Case lStr("Category")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDPT"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDPT"
Case lStr("FTE#")
'    clpOldValue.LookupType = HRTABL
'    clpOldValue.TABLName = "EDAB"
'    clpNewValue.LookupType = HRTABL
'    clpNewValue.TABLName = "EDAB"
    clpNewValue.Visible = False
    clpOldValue.Visible = False
    txtNewValue.Visible = True
    txtOldValue.Visible = True
    
Case lStr("FTE# Hours")
'    clpOldValue.LookupType = HRTABL
'    clpOldValue.TABLName = "EDAB"
'    clpNewValue.LookupType = HRTABL
'    clpNewValue.TABLName = "EDAB"
    clpNewValue.Visible = False
    clpOldValue.Visible = False
    txtNewValue.Visible = True
    txtOldValue.Visible = True
Case lStr("Location")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDLC"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDLC"
Case lStr("Union")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDOR"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDOR"
Case lStr("Region")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDRG"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDRG"
'Hemu - Begin - Uncommented
Case lStr("Salary")
    clpOldValue.Enabled = False 'Hemu
    clpNewValue.Enabled = False 'Hemu
    txtNewValue.Visible = False 'Hemu
    txtOldValue.Visible = False 'Hemu
    lblNewValue.Visible = False 'Hemu
    lblOldValue.Visible = False 'Hemu
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        chkChgSalary.Value = 1
    End If

'    lblTitle(5).Visible = True
'    lblTitle(6).Visible = True
'    medsalary.Visible = True
'    comPayPer.Visible = True
'Hemu - End
Case lStr("Section")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDSE"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDSE"
Case lStr("Status")
    clpNewValue.Visible = True
    clpOldValue.Visible = True
    
    txtNewValue.Visible = False
    txtOldValue.Visible = False

    clpOldValue.LookupType = HRTABL
    clpOldValue.TablName = "EDEM"
    clpNewValue.LookupType = HRTABL
    clpNewValue.TablName = "EDEM"
Case ("Smoker") 'Ticket #21118 Franks 10/28/2011
    clpOldValue.Visible = False
    clpNewValue.Visible = False
    txtNewValue.Visible = True
    txtOldValue.Visible = True
Case ("Marital Status") 'Ticket #21118 Franks 10/28/2011
    clpOldValue.Visible = False
    clpNewValue.Visible = False
    txtNewValue.Visible = True
    txtOldValue.Visible = True
End Select
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HREMPHIS", "SELECT")

End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_EmpHis_ViewOwn Then
        MsgBox "You cannot view your own Employee History.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

Select Case gsSystemDb
Case "MS SQL SERVER"
    SQLQ = "Select *, "
    'TYPE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
    If glbWFC Then 'Ticket #21118 Franks 10/28/2011
        SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN '" & ("Smoker") & "' "
        'Ticket #28794 - Opening up Marital Status for everyone
        'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
    End If
    'Ticket #28794 - Opening up Marital Status for everyone
    SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
    
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN '" & "Position" & "' " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN '" & lStr("Rept. Authority 1") & "' " 'Ticket #27553 Franks 09/22/2015
    'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
    SQLQ = SQLQ & " END) AS HISTYPE, "
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        'SALARY
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN cast(EE_SALARY AS varchar(20)) "
        SQLQ = SQLQ & " ELSE '' END) AS SALARY, "
        
        'PER
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
        SQLQ = SQLQ & " ELSE '' END) AS SALCD, "
    End If
    
    'OLDVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_OLDFTE AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_OLDFTEHR AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
    If glbWFC Then '#21118 Franks 10/28/2011
        'SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN (CASE WHEN EE_OLDSMOKER = 0 THEN 'No' WHEN EE_OLDSMOKER = 1 THEN 'Yes' END) "
        SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_OLDSMOKER "
        'Ticket #28794 - Opening up Marital Status for everyone
        'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
    End If
    'Ticket #28794 - Opening up Marital Status for everyone
    SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
    
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_OLDPOSITION " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_OLDREPORT1 " 'Ticket #27553 Franks 09/21/2015
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    End If
    SQLQ = SQLQ & " END) AS OLDVALUE, "
    
    'NEWVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_NEWFTE AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_NEWFTEHR AS varchar(20)) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
    If glbWFC Then '#21118 Franks 10/28/2011
        'SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN (CASE WHEN EE_NEWSMOKER =0 THEN 'No' WHEN EE_NEWSMOKER = 1 THEN 'Yes' END) "
        SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_NEWSMOKER  "
        'Ticket #28794 - Opening up Marital Status for everyone
        'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
    End If
    'Ticket #28794 - Opening up Marital Status for everyone
    SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
    
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_NEWPOSITION " 'Ticket #27553 Franks 09/21/2015
    SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_NEWREPORT1 " 'Ticket #27553 Franks 09/21/2015
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    End If
    SQLQ = SQLQ & " END) AS NEWVALUE "
Case "ORACLE"
    If glbtermopen Then
        SQLQ = "Select Term_HREMPHIS.*, "
    Else
        SQLQ = "Select HREMPHIS.*, "
    End If
    
    'TYPE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN '" & "Position" & "' " 'Ticket #27553 Franks 09/21/2015
    'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
    SQLQ = SQLQ & " END) AS HISTYPE, "
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        'SALARY
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN TO_CHAR(EE_SALARY ) "
        SQLQ = SQLQ & " ELSE '' END) AS SALARY, "
        'PER
        SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
        SQLQ = SQLQ & " ELSE '' END) AS SALCD, "
    End If
    
    'OLDVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_OLDFTE) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_OLDFTEHR) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_OLDPOSITION " 'Ticket #27553 Franks 09/21/2015

    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    End If
    SQLQ = SQLQ & " END) AS OLDVALUE, "
    
    'NEWVALUE
    SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
    SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
    SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
    SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
    SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
    SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_NEWFTE) "
    SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_NEWFTEHR) "
    SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
    SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
    SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
    SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
    SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
    SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_NEWPOSITION " 'Ticket #27553 Franks 09/21/2015
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
    End If
    SQLQ = SQLQ & " END) AS NEWVALUE "
Case Else
    SQLQ = "Select *, "
    SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , 'Department ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , 'Division ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , 'Status ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , 'FT/PT/SE/TR/OT ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , 'Union ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , 'FTE# ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , 'FTE# Hours ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , 'Region ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , 'Section ',"
    SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , 'Administered By ' ,"
    SQLQ = SQLQ & " IIF(EE_NEWLOC IS NOT NULL , 'Location '  ,"
    SQLQ = SQLQ & " IIF(EE_NEWBENEGROUP IS NOT NULL , 'Benefit Group', '')"
    SQLQ = SQLQ & " )))))))))))"
    SQLQ = SQLQ & " AS HISTYPE,"
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , str(EE_SALARY), '') AS SALARY,"
        SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , EE_SALCD , '') AS SALCD,"
    End If
    SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , EE_OLDDEPT ,"
    SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , EE_OLDDIV ,"
    SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , EE_OLDSTAT ,"
    SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , EE_OLDPT ,"
    SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , EE_OLDORG,"
    SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , str(EE_OLDFTE ),"
    SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , str(EE_OLDFTEHR),"
    SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , EE_OLDREGION,"
    SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , EE_OLDSECTION,"
    SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_OLDADMINBY ,"
    SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_OLDLOC ,"
    SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_OLDBENEGROUP ,"
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , ''  ,'')"
    Else
        SQLQ = SQLQ & "'')"
    End If
    SQLQ = SQLQ & " ))))))))))))"
    SQLQ = SQLQ & "  AS OLDVALUE,"
    
    SQLQ = SQLQ & " IIF( EE_NEWDEPT IS NOT NULL , EE_NEWDEPT  ,"
    SQLQ = SQLQ & " IIF( EE_NEWDIV IS NOT NULL , EE_NEWDIV  ,"
    SQLQ = SQLQ & " IIF( EE_NEWSTAT IS NOT NULL , EE_NEWSTAT  ,"
    SQLQ = SQLQ & " IIF( EE_NEWPT IS NOT NULL , EE_NEWPT  ,"
    SQLQ = SQLQ & " IIF( EE_NEWORG IS NOT NULL , EE_NEWORG  ,"
    SQLQ = SQLQ & " IIF( EE_NEWFTE IS NOT NULL , str(EE_NEWFTE),"
    SQLQ = SQLQ & " IIF( EE_NEWFTEHR IS NOT NULL , str(EE_NEWFTEHR),"
    SQLQ = SQLQ & " IIF( EE_NEWREGION IS NOT NULL , EE_NEWREGION  ,"
    SQLQ = SQLQ & " IIF( EE_NEWSECTION IS NOT NULL , EE_NEWSECTION,"
    SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_NEWADMINBY  ,"
    SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_NEWLOC  ,"
    SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_NEWBENEGROUP  ,"
    
    'Ticket #14963 - Check if the user has access to this employee's salary information
    If gSec_Inq_Salary Then
        SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , '' ,'' )"
    Else
        SQLQ = SQLQ & "'')"
    End If
    SQLQ = SQLQ & " ))))))))))))"
    SQLQ = SQLQ & " AS NEWVALUE"
End Select
If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_HREMPHIS "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = SQLQ & " FROM HREMPHIS "
    SQLQ = SQLQ & " WHERE EE_EMPNBR = " & glbLEE_ID
End If
fglbSQL = SQLQ
SQLQ = SQLQ & " ORDER BY EE_CHGDATE DESC, EE_ID DESC"

Data1.RecordSource = SQLQ
Data1.Refresh
'Set FRS = Data1.Recordset.Clone


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HREMPHIS", "SELECT")
Resume Next

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "frmEHistory"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmEHistory"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "frmEHistory"
If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS


Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_EmpHis_ViewOwn Then
        MsgBox "You cannot view your own Employee History.", vbCritical, "info:HR - Security"
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
'vbxTrueGrid.FetchRowStyle = True
'vbxTrueGrid.MarqueeStyle = 3

If glbWFC Then 'Ticket #21119 Franks 11/14/2011
    Call WFCNGSSubGroup
End If

If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Employee History - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value
Call ST_UPD_MODE(False)
If Not gSec_Upd_Basic Then
'    cmdNew.Enabled = False
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

'Ticket #18435 Frank 04/29/2010
'Samuel doesn't want to see salary on this screen. Jerry asked to hide all salary fields
'Ticket #14963 - Check if the user has access to this employee's salary information
'chkChgSalary.Visible = gSec_Inq_Salary
'lblTitle(5).Visible = gSec_Inq_Salary
'medsalary.Visible = gSec_Inq_Salary
'lblTitle(6).Visible = gSec_Inq_Salary
'comPayPer.Visible = gSec_Inq_Salary

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
comType.Clear
comType.AddItem lStr("Administered By")
comType.AddItem lStr("Benefit Group")
comType.AddItem lStr("Department")
comType.AddItem lStr("Division")
comType.AddItem lStr("Category")
comType.AddItem lStr("FTE#")
comType.AddItem lStr("FTE# Hours")
comType.AddItem lStr("Location")
comType.AddItem lStr("Union")
comType.AddItem lStr("Region")
'comType.AddItem lStr("Salary")
comType.AddItem lStr("Section")
comType.AddItem lStr("Status")
If glbWFC Then 'Ticket #21118 Franks 10/28/2011
    comType.AddItem lStr("Smoker")
    'Ticket #28794 - Openining up Marital Status for everyone
    'comType.AddItem lStr("Marital Status")
    'Call ComMStat
End If
'Ticket #28794 - Openining up Marital Status for everyone
comType.AddItem lStr("Marital Status")
Call ComMStat

comType.AddItem "Position" 'Ticket #27553 Franks 09/21/2015
comType.AddItem lStr("Rept. Authority 1") 'Ticket #27553 Franks 09/21/2015
comPayPer.Clear
comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour "
comPayPer.AddItem "Monthly "

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT

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
    Call NextForm
End Sub

Private Sub lblDesc_Click()

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

fUPMode = TF    ' update mode

 
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
End If
clpNewValue.Enabled = TF
'If comType = "Salary" Then clpNewValue.Enabled = False

comType.Enabled = TF
dlpCHGDate.Enabled = TF
'vbxTrueGrid.Enabled = FT

End Sub

Private Sub lblSalCode_Change()
Dim x
For x = 0 To comPayPer.ListCount - 1
    If UCase(lblSalCode) = Left(UCase(comPayPer.List(x)), 1) Then
        comPayPer.ListIndex = x
    End If
Next
End Sub

Private Sub lblType_Change()
Dim x
For x = 0 To comType.ListCount - 1
    If UCase(Trim(lblType)) = UCase(Trim(comType.List(x))) Then
        comType.ListIndex = x
    End If
Next

End Sub

Private Sub txtNewValue_Change()
    If clpNewValue.Text <> txtNewValue.Text Then clpNewValue.Text = txtNewValue.Text
    
End Sub

Private Sub txtOldValue_Change()
    If Not clpOldValue.Text = txtOldValue.Text Then clpOldValue.Text = txtOldValue.Text
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
On Error GoTo Eh
    'added by Bryan 18/Jan/06 Ticket#10222
'    FRS.Requery
'    FRS.Bookmark = Bookmark
    'change row colour
'    If FRS("BD_FREEZE") = True Then
'        RowStyle.ForeColor = vbRed
'    End If
    
Eh:
    Exit Sub
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
        
        Select Case gsSystemDb
        Case "MS SQL SERVER"
            SQLQ = "Select *, "
            
            'TYPE
            SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
            If glbWFC Then 'Ticket #21118 Franks 10/28/2011
                SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN '" & ("Smoker") & "' "
                'Ticket #28794 - Opening up Marital Status for everyone
                'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
            End If
            'Ticket #28794 - Opening up Marital Status for everyone
            SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN '" & ("Marital Status") & "' "
            
            SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN '" & "Position" & "' " 'Ticket #27553 Franks 09/21/2015
            SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN '" & lStr("Rept. Authority 1") & "' " 'Ticket #27553 Franks 09/22/2015
            
            'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
            SQLQ = SQLQ & " END) AS HISTYPE, "
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                'SALARY
                SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN cast(EE_SALARY AS varchar(20)) "
                SQLQ = SQLQ & " ELSE '' END) AS SALARY, "
                'PER
                SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
                SQLQ = SQLQ & " ELSE '' END) AS SALCD, "
            End If
            
            'OLDVALUE
            SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
            SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
            SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
            SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
            SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
            SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_OLDFTE AS varchar(20)) "
            SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_OLDFTEHR AS varchar(20)) "
            SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
            SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
            SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
            SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
            SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
            If glbWFC Then '#21118 Franks 10/28/2011
                'SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN (CASE WHEN EE_OLDSMOKER = 0 THEN 'No' WHEN EE_OLDSMOKER = 1 THEN 'Yes' END) "
                SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_OLDSMOKER "
                'Ticket #28794 - Opening up Marital Status for everyone
                'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
            End If
            'Ticket #28794 - Opening up Marital Status for everyone
            SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_OLDMSTAT "
            
            SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_OLDPOSITION " 'Ticket #27553 Franks 09/21/2015
            SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_OLDREPORT1 " 'Ticket #27553 Franks 09/21/2015
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
            End If
            
            SQLQ = SQLQ & " END) AS OLDVALUE, "
            
            'NEWVALUE
            SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
            SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
            SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
            SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
            SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
            SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN cast(EE_NEWFTE AS varchar(20)) "
            SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN cast(EE_NEWFTEHR AS varchar(20)) "
            SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
            SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
            SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
            SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
            SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
            If glbWFC Then '#21118 Franks 10/28/2011
                'SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN (CASE WHEN EE_NEWSMOKER =0 THEN 'No' WHEN EE_NEWSMOKER = 1 THEN 'Yes' END) "
                SQLQ = SQLQ & " WHEN EE_NEWSMOKER IS NOT NULL THEN EE_NEWSMOKER  "
                'Ticket #28794 - Opening up Marital Status for everyone
                'SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
            End If
            'Ticket #28794 - Opening up Marital Status for everyone
            SQLQ = SQLQ & " WHEN EE_NEWMSTAT IS NOT NULL THEN EE_NEWMSTAT "
            
            SQLQ = SQLQ & " WHEN EE_NEWPOSITION IS NOT NULL THEN EE_NEWPOSITION " 'Ticket #27553 Franks 09/21/2015
            SQLQ = SQLQ & " WHEN EE_NEWREPORT1 IS NOT NULL THEN EE_NEWREPORT1 " 'Ticket #27553 Franks 09/21/2015
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
            End If
            SQLQ = SQLQ & " END) AS NEWVALUE "
            
        Case "ORACLE"
            If glbtermopen Then
                SQLQ = "Select Term_HREMPHIS.*, "
            Else
                SQLQ = "Select HREMPHIS.*, "
            End If
            
            'TYPE
            SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN '" & lStr("Department") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN '" & lStr("Division") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN '" & lStr("Status") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN '" & lStr("Category") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN '" & lStr("Union") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN '" & lStr("FTE#") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN '" & lStr("FTE# Hours") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN '" & lStr("Region") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN '" & lStr("Section") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN '" & lStr("Administered By") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN '" & lStr("Location") & " ' "
            SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN '" & lStr("Benefit Group") & " ' "
            'SQLQ = SQLQ & " WHEN EE_SALARY <> 0 THEN 'Salary' "
            SQLQ = SQLQ & " END) AS HISTYPE, "
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                'SALARY
                SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN TO_CHAR(EE_SALARY ) "
                SQLQ = SQLQ & " ELSE '' END) AS SALARY, "
                'PER
                SQLQ = SQLQ & " (CASE WHEN EE_SALARY <> 0 THEN EE_SALCD "
                SQLQ = SQLQ & " ELSE '' END) AS SALCD, "
            End If
            
            'OLDVALUE
            SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_OLDDEPT "
            SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_OLDDIV "
            SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_OLDSTAT "
            SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_OLDPT "
            SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_OLDORG "
            SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_OLDFTE) "
            SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_OLDFTEHR) "
            SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_OLDREGION "
            SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_OLDSECTION "
            SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_OLDADMINBY "
            SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_OLDLOC "
            SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_OLDBENEGROUP "
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
            End If
            SQLQ = SQLQ & " END) AS OLDVALUE, "
            
            'NEWVALUE
            SQLQ = SQLQ & " (CASE WHEN EE_NEWDEPT IS NOT NULL THEN EE_NEWDEPT "
            SQLQ = SQLQ & " WHEN EE_NEWDIV IS NOT NULL THEN EE_NEWDIV "
            SQLQ = SQLQ & " WHEN EE_NEWSTAT IS NOT NULL THEN EE_NEWSTAT "
            SQLQ = SQLQ & " WHEN EE_NEWPT IS NOT NULL THEN EE_NEWPT "
            SQLQ = SQLQ & " WHEN EE_NEWORG IS NOT NULL THEN EE_NEWORG "
            SQLQ = SQLQ & " WHEN EE_NEWFTE IS NOT NULL THEN TO_CHAR(EE_NEWFTE) "
            SQLQ = SQLQ & " WHEN EE_NEWFTEHR IS NOT NULL THEN TO_CHAR(EE_NEWFTEHR) "
            SQLQ = SQLQ & " WHEN EE_NEWREGION IS NOT NULL THEN EE_NEWREGION "
            SQLQ = SQLQ & " WHEN EE_NEWSECTION IS NOT NULL THEN EE_NEWSECTION "
            SQLQ = SQLQ & " WHEN EE_NEWADMINBY IS NOT NULL THEN EE_NEWADMINBY "
            SQLQ = SQLQ & " WHEN EE_NEWLOC IS NOT NULL THEN EE_NEWLOC "
            SQLQ = SQLQ & " WHEN EE_NEWBENEGROUP IS NOT NULL THEN EE_NEWBENEGROUP "
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " WHEN EE_SALARY IS NOT NULL THEN '' "
            End If
            SQLQ = SQLQ & " END) AS NEWVALUE "
        Case Else
            SQLQ = "Select *, "
            SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , 'Department ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , 'Division ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , 'Status ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , 'FT/PT/SE/TR/OT ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , 'Union ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , 'FTE# ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , 'FTE# Hours ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , 'Region ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , 'Section ',"
            SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , 'Administered By ' ,"
            SQLQ = SQLQ & " IIF(EE_NEWLOC IS NOT NULL , 'Location '  ,"
            SQLQ = SQLQ & " IIF(EE_NEWBENEGROUP IS NOT NULL , 'Benefit Group', '')"
            SQLQ = SQLQ & " )))))))))))"
            SQLQ = SQLQ & " AS HISTYPE,"
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , str(EE_SALARY), '') AS SALARY,"
                SQLQ = SQLQ & " IIF( EE_SALARY <> 0 , EE_SALCD , '') AS SALCD,"
            End If
            
            SQLQ = SQLQ & " IIF(EE_NEWDEPT IS NOT NULL , EE_OLDDEPT ,"
            SQLQ = SQLQ & " IIF(EE_NEWDIV IS NOT NULL , EE_OLDDIV ,"
            SQLQ = SQLQ & " IIF(EE_NEWSTAT IS NOT NULL , EE_OLDSTAT ,"
            SQLQ = SQLQ & " IIF(EE_NEWPT IS NOT NULL , EE_OLDPT ,"
            SQLQ = SQLQ & " IIF(EE_NEWORG IS NOT NULL , EE_OLDORG,"
            SQLQ = SQLQ & " IIF(EE_NEWFTE IS NOT NULL , str(EE_OLDFTE ),"
            SQLQ = SQLQ & " IIF(EE_NEWFTEHR IS NOT NULL , str(EE_OLDFTEHR),"
            SQLQ = SQLQ & " IIF(EE_NEWREGION IS NOT NULL , EE_OLDREGION,"
            SQLQ = SQLQ & " IIF(EE_NEWSECTION IS NOT NULL , EE_OLDSECTION,"
            SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_OLDADMINBY ,"
            SQLQ = SQLQ & " IIF(EE_NEWLOC IS NOT NULL , EE_OLDLOC ,"
            SQLQ = SQLQ & " IIF(EE_NEWBENEGROUP IS NOT NULL , EE_OLDBENEGROUP ,"
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " IIF(EE_SALARY IS NOT NULL , ''  ,'')"
            Else
                SQLQ = SQLQ & "'')"
            End If
            
            SQLQ = SQLQ & " ))))))))))))"
            SQLQ = SQLQ & "  AS OLDVALUE,"
            
            SQLQ = SQLQ & " IIF( EE_NEWDEPT IS NOT NULL , EE_NEWDEPT  ,"
            SQLQ = SQLQ & " IIF( EE_NEWDIV IS NOT NULL , EE_NEWDIV  ,"
            SQLQ = SQLQ & " IIF( EE_NEWSTAT IS NOT NULL , EE_NEWSTAT  ,"
            SQLQ = SQLQ & " IIF( EE_NEWPT IS NOT NULL , EE_NEWPT  ,"
            SQLQ = SQLQ & " IIF( EE_NEWORG IS NOT NULL , EE_NEWORG  ,"
            SQLQ = SQLQ & " IIF( EE_NEWFTE IS NOT NULL , str(EE_NEWFTE),"
            SQLQ = SQLQ & " IIF( EE_NEWFTEHR IS NOT NULL , str(EE_NEWFTEHR),"
            SQLQ = SQLQ & " IIF( EE_NEWREGION IS NOT NULL , EE_NEWREGION  ,"
            SQLQ = SQLQ & " IIF( EE_NEWSECTION IS NOT NULL , EE_NEWSECTION,"
            SQLQ = SQLQ & " IIF(EE_NEWADMINBY IS NOT NULL , EE_NEWADMINBY  ,"
            SQLQ = SQLQ & " IIF( EE_NEWLOC IS NOT NULL , EE_NEWLOC  ,"
            SQLQ = SQLQ & " IIF( EE_NEWBENEGROUP IS NOT NULL , EE_NEWBENEGROUP  ,"
            
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                SQLQ = SQLQ & " IIF( EE_SALARY IS NOT NULL , '' ,'' )"
            Else
                SQLQ = SQLQ & "'')"
            End If
            
            SQLQ = SQLQ & " ))))))))))))"
            SQLQ = SQLQ & " AS NEWVALUE"
        End Select
        If glbtermopen Then
            SQLQ = SQLQ & " FROM Term_HREMPHIS "
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = SQLQ & " FROM HREMPHIS "
            SQLQ = SQLQ & " WHERE EE_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Values As String, Valuesvl As String, ValuesDte As String
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err

Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HREMPHIS", "Add")
Resume Next

End Sub

Private Function RollBack()
On Error GoTo RR
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
RR:
End Function

Sub Display_Value()
Dim SQLQ
On Error Resume Next
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    Me.clpNewValue.Enabled = False
    Me.clpOldValue.Enabled = False
    Me.chkChgSalary.Enabled = False
    Me.dlpCHGDate.Enabled = False
    comType.Enabled = False
Else
    txtNewValue.Visible = False
    txtOldValue.Visible = False
    Select Case Trim(lStr(vbxTrueGrid.Columns(1).Value))
        Case lStr("Administered By")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
        
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDAB"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDAB"
        Case lStr("Benefit Group")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "BGMF"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "BGMF"
        Case lStr("Department")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = Department
            clpNewValue.LookupType = Department
        Case lStr("Division")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = Division
            clpNewValue.LookupType = Division
        Case lStr("Category")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDPT"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDPT"
        Case lStr("FTE#")
        '    clpOldValue.LookupType = HRTABL
        '    clpOldValue.TABLName = "EDAB"
        '    clpNewValue.LookupType = HRTABL
        '    clpNewValue.TABLName = "EDAB"
            clpNewValue.Visible = False
            clpOldValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
        Case lStr("FTE# Hours")
        '    clpOldValue.LookupType = HRTABL
        '    clpOldValue.TABLName = "EDAB"
        '    clpNewValue.LookupType = HRTABL
        '    clpNewValue.TABLName = "EDAB"
            clpNewValue.Visible = False
            clpOldValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
        Case lStr("Location")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDLC"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDLC"
        Case lStr("Union")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDOR"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDOR"
        Case lStr("Region")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDRG"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDRG"
        Case lStr("Salary")
            clpOldValue.Enabled = False 'Hemu
            clpNewValue.Enabled = False 'Hemu
            txtNewValue.Visible = False 'Hemu
            txtOldValue.Visible = False 'Hemu
            lblNewValue.Visible = False 'Hemu
            lblOldValue.Visible = False 'Hemu
        
        '    clpOldValue.Enabled = False
        '    clpNewValue.Enabled = False
        '    lblTitle(5).Visible = True
        '    lblTitle(6).Visible = True
        '    medsalary.Visible = True
        '    comPayPer.Visible = True
        Case lStr("Section")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDSE"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDSE"
        Case lStr("Status")
            clpNewValue.Visible = True
            clpOldValue.Visible = True
            txtNewValue.Visible = False
            txtOldValue.Visible = False
            
            clpOldValue.LookupType = HRTABL
            clpOldValue.TablName = "EDEM"
            clpNewValue.LookupType = HRTABL
            clpNewValue.TablName = "EDEM"
        Case ("Smoker") 'Ticket #21118 Franks 10/28/2011
            clpOldValue.Visible = False
            clpNewValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
        Case ("Marital Status") 'Ticket #21118 Franks 10/28/2011
            clpOldValue.Visible = False
            clpNewValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
        Case ("Position") 'Ticket #27553 Franks 09/21/2015
            clpOldValue.Visible = False
            clpNewValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
        Case lStr("Rept. Authority 1") 'Ticket #27553 Franks 09/21/2015
            clpOldValue.Visible = False
            clpNewValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
    End Select
    
    '''Ticket #21118 Franks 10/28/2011
    ''If Trim(lStr(vbxTrueGrid.Columns(1).Value)) = ("Marital Status") Then
    ''    ComMStatOld.Visible = True
    ''    ComMStatNew.Visible = True
    ''    txtOldValue.Visible = False
    ''    txtNewValue.Visible = False
    ''Else
    ''    ComMStatOld.Visible = False
    ''    ComMStatNew.Visible = False
    ''End If
End If

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    If glbtermopen Then
        RSDATA.Open fglbSQL, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        RSDATA.Open fglbSQL, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    lblTitle(5).Visible = False
    lblTitle(6).Visible = False
    medsalary.Visible = False
    comPayPer.Visible = False
Else
    If glbtermopen Then
        SQLQ = fglbSQL
        SQLQ = SQLQ & " AND EE_ID = " & Data1.Recordset!EE_ID
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = fglbSQL
        SQLQ = SQLQ & " AND EE_ID = " & Data1.Recordset!EE_ID
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    If IsNull(RSDATA("EE_SALARY")) Then
        comType.Visible = True
        lbType.Visible = True
        lblNewValue.Visible = True
        lblOldValue.Visible = True
    
        lblTitle(5).Visible = False
        lblTitle(6).Visible = False
        medsalary.Visible = False
        comPayPer.Visible = False
    
        'Ticket #14963 - Check if the user has access to this employee's salary information
        If gSec_Inq_Salary Then
            chkChgSalary = 0
        End If
    ElseIf RSDATA("EE_SALARY") = 0 Then
        comType.Visible = True
        lbType.Visible = True
        lblNewValue.Visible = True
        lblOldValue.Visible = True
        
        lblTitle(5).Visible = False
        lblTitle(6).Visible = False
        medsalary.Visible = False
        comPayPer.Visible = False
        
        'Ticket #14963 - Check if the user has access to this employee's salary information
        If gSec_Inq_Salary Then
            chkChgSalary = 0
        End If
    Else
        If Trim(RSDATA("HISTYPE")) = "Salary" Or Trim(RSDATA("HISTYPE")) = "" Or IsNull(RSDATA("HISTYPE")) Then
            'Ticket #14963 - Check if the user has access to this employee's salary information
            If gSec_Inq_Salary Then
                chkChgSalary = 1
            End If
        Else
            comType.Visible = True
            lbType.Visible = True
            lblNewValue.Visible = True
            lblOldValue.Visible = True
            
            'Ticket #18435 Frank 04/29/2010
            'Samuel doesn't want to see salary on this screen. Jerry asked to hide all salary fields
            ''Ticket #14963 - Check if the user has access to this employee's salary information
            'If gSec_Inq_Salary Then
            '    lblTitle(5).Visible = True
            '    lblTitle(6).Visible = True
            '    medsalary.Visible = True
            '    comPayPer.Visible = True
            'End If
        End If
    End If
    If IsNull(RSDATA("HISTYPE")) Then
    ElseIf Trim(RSDATA("HISTYPE")) = lStr("FTE#") Or Trim(RSDATA("HISTYPE")) = lStr("FTE# Hours") Then
            clpNewValue.Visible = False
            clpOldValue.Visible = False
            txtNewValue.Visible = True
            txtOldValue.Visible = True
            txtNewValue.Text = RSDATA("Newvalue")
            txtOldValue.Text = RSDATA("oldvalue")
    End If
    
    'Ticket #12708 - Begin
    If clpOldValue.TablName = "EDRG" Then
        If Len(clpOldValue.Text) > 0 Then
            If Not IsNull(RSDATA("EE_OLDREGIONDESC")) Then
                If Len(RSDATA("EE_OLDREGIONDESC")) > 0 Then
                    clpOldValue.Caption = RSDATA("EE_OLDREGIONDESC")
                End If
            End If
        End If
    End If
    If clpNewValue.TablName = "EDRG" Then
        If Len(clpNewValue.Text) > 0 Then
            If Not IsNull(RSDATA("EE_NEWREGIONDESC")) Then
                If Len(RSDATA("EE_NEWREGIONDESC")) > 0 Then
                    clpNewValue.Caption = RSDATA("EE_NEWREGIONDESC")
                End If
            End If
        End If
    End If
    'Ticket #12708 - End
    
    'Ticket #21118 Franks 10/28/2011 - begin
    If Not IsNull(RSDATA("HISTYPE")) Then
        If RSDATA("HISTYPE") = "Smoker" Then
            txtOldValue.Text = RSDATA("EE_OLDSMOKER") 'rsDATA("OLDVALUE")
            txtNewValue.Text = RSDATA("EE_NEWSMOKER") 'rsDATA("NEWVALUE")
        End If
        If Trim(RSDATA("HISTYPE")) = "Marital Status" Then
            txtOldValue.Text = RSDATA("EE_OLDMSTAT")
            txtNewValue.Text = RSDATA("EE_NEWMSTAT")
        End If
    End If
    If glbWFC Then 'Ticket #21119 Franks 11/14/2011

    End If
    'Ticket #21118 Franks 10/28/2011 - end
    
    'Ticket #27553 Franks 09/21/2015 - begin
    If Not IsNull(RSDATA("HISTYPE")) Then
        If RSDATA("HISTYPE") = "Position" Then
            txtOldValue.Text = RSDATA("EE_OLDPOSITION") 'rsDATA("OLDVALUE")
            txtNewValue.Text = RSDATA("EE_NEWPOSITION") 'rsDATA("NEWVALUE")
        End If
        If RSDATA("HISTYPE") = lStr("Rept. Authority 1") Then
            txtOldValue.Text = RSDATA("EE_OLDREPORT1") 'rsDATA("OLDVALUE")
            txtNewValue.Text = RSDATA("EE_NEWREPORT1") 'rsDATA("NEWVALUE")
        End If
    End If
    'Ticket #27553 Franks 09/21/2015 - end
End If
Call SET_UP_MODE
End Sub

Private Sub Setup4WFC_Smoker()
    If Len(glbWFCNGSSubGroup) > 0 Then 'NGS only
    
    End If
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
UpdateRight = gSec_Upd_EMP_HISTORY
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
ElseIf RSDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEHistory.Caption = "Employee History - " & Left$(glbLEE_SName, 5)
    frmEHistory.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub ComMStat() 'Ticket #21118 Franks 10/28/2011
ComMStatOld.Clear
ComMStatOld.AddItem "Single"
ComMStatOld.AddItem "Married"
ComMStatOld.AddItem "Family"
ComMStatOld.AddItem "Parent(Single)"
ComMStatOld.AddItem "Divorced"
ComMStatOld.AddItem "Widowed"
ComMStatOld.AddItem "Common-Law"
ComMStatOld.AddItem "Partner"
ComMStatOld.AddItem "Same-Sex"
ComMStatOld.AddItem "Other"
ComMStatOld.AddItem "Separated"
ComMStatOld.ListIndex = 0

ComMStatNew.Clear
ComMStatNew.AddItem "Single"
ComMStatNew.AddItem "Married"
ComMStatNew.AddItem "Family"
ComMStatNew.AddItem "Parent(Single)"
ComMStatNew.AddItem "Divorced"
ComMStatNew.AddItem "Widowed"
ComMStatNew.AddItem "Common-Law"
ComMStatNew.AddItem "Partner"
ComMStatNew.AddItem "Same-Sex"
ComMStatNew.AddItem "Other"
ComMStatNew.AddItem "Separated"
ComMStatNew.ListIndex = 0

If glbCompSerial = "S/N - 2482W" Then 'Ticket #28794 - Windsor Family Credit Union -
    ComMStatOld.Clear
    ComMStatOld.AddItem "Single"
    ComMStatOld.AddItem "Married"
    ComMStatOld.AddItem "Common-Law"
    ComMStatOld.AddItem "Widow/Widower"
    ComMStatOld.AddItem "Separated"
    ComMStatOld.AddItem "Divorced"
    ComMStatOld.ListIndex = 0
    
    ComMStatNew.Clear
    ComMStatNew.AddItem "Single"
    ComMStatNew.AddItem "Married"
    ComMStatNew.AddItem "Common-Law"
    ComMStatNew.AddItem "Widow/Widower"
    ComMStatNew.AddItem "Separated"
    ComMStatNew.AddItem "Divorced"
    ComMStatNew.ListIndex = 0
End If

ComMStatOld.Top = txtOldValue.Top
ComMStatOld.Left = txtOldValue.Left
ComMStatNew.Top = txtNewValue.Top
ComMStatNew.Left = txtNewValue.Left
End Sub

Private Function GetComMStatInx(xVal) 'Ticket #21118 Franks 10/28/2011
Dim retVal
retVal = -1
If xVal = "S" Then retVal = 0
If xVal = "M" Then retVal = 1
If xVal = "F" Then retVal = 2
If xVal = "P" Then retVal = 3
If xVal = "D" Then retVal = 4
If xVal = "W" Then retVal = 5
If xVal = "C" Then retVal = 6
If xVal = "R" Then retVal = 7
If xVal = "X" Then retVal = 8
If xVal = "O" Then retVal = 9
If xVal = "A" Then retVal = 10

If glbCompSerial = "S/N - 2482W" Then 'Ticket #28794 - Windsor Family Credit Union
    retVal = -1
    If xVal = "S" Then retVal = 0
    If xVal = "M" Then retVal = 1
    If xVal = "C" Then retVal = 2
    If xVal = "W" Then retVal = 3
    If xVal = "A" Then retVal = 4
    If xVal = "D" Then retVal = 5
End If

GetComMStatInx = retVal
End Function

Private Sub WFCNGSSubGroup() 'Ticket #21119 Franks 11/14/2011
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String
    glbWFCNGSSubGroup = ""
    If glbtermopen Then
        SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM Term_HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & glbTERM_ID & " "
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    End If
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
    End If
    rsEmpee.Close
End Sub

