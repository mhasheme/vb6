VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000F&
   Caption         =   "info:HR"
   ClientHeight    =   8655
   ClientLeft      =   270
   ClientTop       =   465
   ClientWidth     =   12840
   Icon            =   "fmdimain.frx":0000
   Begin Threed.SSPanel lstPanel 
      Align           =   3  'Align Left
      Height          =   7875
      Left            =   3600
      TabIndex        =   10
      Top             =   360
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   13891
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      Alignment       =   0
      Autosize        =   1
      Begin MSComctlLib.ListView lstView 
         Height          =   7935
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   13996
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         AllowReorder    =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   4
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
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
         OLEDragMode     =   1
         NumItems        =   0
         Picture         =   "fmdimain.frx":000C
      End
   End
   Begin Threed.SSPanel panMain 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   8235
      Width           =   12840
      _Version        =   65536
      _ExtentX        =   22648
      _ExtentY        =   741
      _StockProps     =   15
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin Threed.SSPanel panHelp 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   90
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         FloodShowPct    =   0   'False
         Alignment       =   0
      End
      Begin Threed.SSPanel Panel3D1 
         Height          =   255
         Left            =   8580
         TabIndex        =   2
         Top             =   90
         Width           =   2595
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         FloodShowPct    =   0   'False
         Font3D          =   1
         Alignment       =   0
         Begin VB.Label lblTime 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   2400
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   59000
         Left            =   9840
         Top             =   0
      End
      Begin Crystal.CrystalReport vbxCommonDlg 
         Left            =   10365
         Top             =   0
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
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSComDlg.CommonDialog vbxCommon 
         Left            =   9420
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         FontSize        =   0
         MaxFileSize     =   256
      End
      Begin Threed.SSPanel panHelp 
         Height          =   255
         Index           =   3
         Left            =   8040
         TabIndex        =   3
         Top             =   90
         Width           =   495
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         Alignment       =   0
      End
      Begin Threed.SSPanel panHelp 
         Height          =   255
         Index           =   2
         Left            =   7500
         TabIndex        =   5
         Top             =   90
         Width           =   495
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   15
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         Alignment       =   0
      End
      Begin Threed.SSPanel panHelp 
         Height          =   255
         Index           =   1
         Left            =   4020
         TabIndex        =   4
         Top             =   90
         Width           =   3435
         _Version        =   65536
         _ExtentX        =   6059
         _ExtentY        =   450
         _StockProps     =   15
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   2
         BevelOuter      =   1
         Alignment       =   0
      End
      Begin Threed.SSPanel panHelp 
         Height          =   255
         Index           =   4
         Left            =   11280
         TabIndex        =   14
         Top             =   90
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   450
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
         BorderWidth     =   2
         BevelOuter      =   1
         Alignment       =   0
         Autosize        =   1
      End
   End
   Begin MSComctlLib.Toolbar MainToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "close"
            Description     =   "close"
            Object.ToolTipText     =   "Close"
            ImageKey        =   "close"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "preview"
            Description     =   "report"
            Object.ToolTipText     =   "view"
            ImageKey        =   "preview"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Description     =   "report"
            Object.ToolTipText     =   "print"
            ImageKey        =   "print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "NewApplicant"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewEmployee"
            Description     =   "NewEmployee"
            Object.ToolTipText     =   "New Employee"
            ImageKey        =   "NewApplicant"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "edit"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewRecord"
            Description     =   "edit"
            Object.ToolTipText     =   "New Record"
            ImageKey        =   "NewRecord"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Description     =   "edit"
            Object.ToolTipText     =   "OK/Save"
            Object.Tag             =   "UPDATE"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancel"
            Description     =   "edit"
            Object.ToolTipText     =   "Cancel"
            Object.Tag             =   "UPDATE"
            ImageKey        =   "cancel"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Description     =   "edit"
            Object.ToolTipText     =   "Delete"
            Object.Tag             =   "UPDATE"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "mass"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "massadd"
            Description     =   "mass"
            Object.ToolTipText     =   "Mass Add"
            ImageKey        =   "massadd"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "massupdate"
            Description     =   "mass"
            Object.ToolTipText     =   "Mass Update"
            ImageKey        =   "massupdate"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "massdelete"
            Description     =   "mass"
            Object.ToolTipText     =   "Mass Delete"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "up"
            Description     =   "move"
            Object.ToolTipText     =   "Row Up"
            Object.Tag             =   "LOOKUP"
            ImageKey        =   "up"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "down"
            Description     =   "move"
            Object.ToolTipText     =   "Row Down"
            Object.Tag             =   "LOOKUP"
            ImageKey        =   "down"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "find"
            Object.ToolTipText     =   "Find an Employee"
            Object.Tag             =   "LOOKUP"
            ImageKey        =   "find"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "findA"
                  Text            =   "Active Employee"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "findTerm"
                  Text            =   "Terminated Employee"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Object.Visible         =   0   'False
                  Key             =   "findCREQ"
                  Text            =   "Find Closed Requisition"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "findJOB"
                  Text            =   "Find Job"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "findPOS"
                  Text            =   "Find Position"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "findDiv"
                  Text            =   "Find Division"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "findCandi"
                  Text            =   "Find Candidate"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "mail"
            Object.ToolTipText     =   "Mail Box"
            Object.Tag             =   "Mail Box"
            ImageKey        =   "mail"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Microsoft Word"
            Object.Tag             =   "Microsoft Words"
            ImageKey        =   "word"
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Microsoft Excel"
            Object.Tag             =   "Microsoft Excel"
            ImageKey        =   "excel"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "question"
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hrsoft"
            Description     =   "HRsoft"
            Object.ToolTipText     =   "HRsoft"
            ImageKey        =   "hrsoft5"
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewContractEmp"
            Description     =   "New Contractor"
            Object.ToolTipText     =   "New Contractor"
            ImageKey        =   "employee"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList imlTools 
         Left            =   7560
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   29
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":AFCE0
               Key             =   "find"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B0132
               Key             =   "massadd"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B0584
               Key             =   "NewApplicant"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B089E
               Key             =   "massnew"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B0CF0
               Key             =   "massupdate"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B1142
               Key             =   "NewRecord"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B1684
               Key             =   "save"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B1BC6
               Key             =   "preview"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B2108
               Key             =   "delete"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B221A
               Key             =   "print"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B232C
               Key             =   "cancel"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B243E
               Key             =   "close"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B2890
               Key             =   "down"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B2CE2
               Key             =   "up"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B3134
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B3586
               Key             =   "word"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B3980
               Key             =   "excel"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B3D86
               Key             =   "powerpoint"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B4188
               Key             =   "pdf"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B4228
               Key             =   "adobe"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B4475
               Key             =   "mail"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B48C7
               Key             =   "question"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B4D19
               Key             =   "hrsoft"
               Object.Tag             =   "HRsoft"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B56EE
               Key             =   "hrsoft2"
               Object.Tag             =   "HRsoft"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B6796
               Key             =   "hrsoft3"
               Object.Tag             =   "HRsoft"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B6E82
               Key             =   "hrsoft5"
               Object.Tag             =   "HRsoft"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":B847E
               Key             =   "employee"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":DD6F4
               Key             =   "contract"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":E51FE
               Key             =   "handshake"
            EndProperty
         EndProperty
      End
   End
   Begin Threed.SSPanel panTree 
      Align           =   3  'Align Left
      DragIcon        =   "fmdimain.frx":10A474
      DragMode        =   1  'Automatic
      Height          =   7875
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   3600
      _Version        =   65536
      _ExtentX        =   6350
      _ExtentY        =   13891
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Alignment       =   0
      Begin VB.CommandButton cmdForLostFocus 
         Caption         =   "For Lost Focus"
         Height          =   510
         Left            =   420
         TabIndex        =   13
         Top             =   -510
         Width           =   1590
      End
      Begin VB.Frame fraMove 
         BorderStyle     =   0  'None
         Height          =   8085
         Left            =   3430
         MousePointer    =   9  'Size W E
         TabIndex        =   9
         Top             =   0
         Width           =   190
      End
      Begin MSComctlLib.ImageList imlLarge 
         Left            =   480
         Top             =   1440
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10A8B6
               Key             =   "setup"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10AD08
               Key             =   "mass"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10B15A
               Key             =   "positions"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10B5AC
               Key             =   "folder"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10B9FE
               Key             =   "find"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10BE50
               Key             =   "applicants"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10C2A2
               Key             =   "reports"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10C6F4
               Key             =   "payweb"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlSmall 
         Left            =   1440
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10CB46
               Key             =   "reports"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10CF9A
               Key             =   "requisitions"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10D3EC
               Key             =   "ihrappt"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10D706
               Key             =   "applicants"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":10DB58
               Key             =   "infohr"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":11546B
               Key             =   "find"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":1158BD
               Key             =   "mass"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":115D0F
               Key             =   "Orequisitions"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":116161
               Key             =   "Crequisitions"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":1165B3
               Key             =   "positions"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":116A05
               Key             =   "setup"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":116E57
               Key             =   "Folder"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "fmdimain.frx":1172A9
               Key             =   "payweb"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwTree 
         Height          =   7695
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   13573
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   531
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlSmall"
         Appearance      =   1
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_F_Close 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnu_Sep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_F_PrintSetup 
         Caption         =   "Printer &Setup"
      End
      Begin VB.Menu mnu_F_Preview 
         Caption         =   "Pre&view"
      End
      Begin VB.Menu mnu_F_Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnu_F_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_F_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu_NewEmployee 
      Caption         =   "&New Employee"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_E_NewRecord 
         Caption         =   "New &Record"
      End
      Begin VB.Menu mnu_E_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnu_E_Cancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnu_E_Delete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnu_Mass 
      Caption         =   "Mass C&hanges"
      Begin VB.Menu mnu_M_Add 
         Caption         =   "Mass Add"
      End
      Begin VB.Menu mnu_M_Update 
         Caption         =   "Mass Update"
      End
      Begin VB.Menu mnu_M_Delete 
         Caption         =   "Mass Delete"
      End
   End
   Begin VB.Menu mnuMove 
      Caption         =   "&Move"
      Begin VB.Menu mnu_M_Up 
         Caption         =   "&Up"
      End
      Begin VB.Menu mnu_M_Down 
         Caption         =   "&Down"
      End
   End
   Begin VB.Menu mnuFind 
      Caption         =   "&Find"
      Begin VB.Menu mnu_F_Employee 
         Caption         =   "Find &Active Employee"
      End
      Begin VB.Menu mnu_F_TEmployee 
         Caption         =   "Find &Terminated Employee"
      End
      Begin VB.Menu mnu_F_JobMaster 
         Caption         =   "Find &Job"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_F_Position 
         Caption         =   "Find &Position"
      End
      Begin VB.Menu mnu_F_Division 
         Caption         =   "Find &Division"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_F_Candidate 
         Caption         =   "Find Candidate"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mmnu_File 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnu_Print 
         Caption         =   "P&rint"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_Printer 
         Caption         =   "&Printer Setup"
      End
      Begin VB.Menu mmnu_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mmnu_Find 
      Caption         =   "&Find"
      Visible         =   0   'False
   End
   Begin VB.Menu mmnu_mail 
      Caption         =   "Emai&l"
   End
   Begin VB.Menu mmnu_Custom 
      Caption         =   "&Custom Features"
   End
   Begin VB.Menu mmnu_ImportExport 
      Caption         =   "&Import/Export"
   End
   Begin VB.Menu mmnu_AppT 
      Caption         =   "&Applicant Tracking"
   End
   Begin VB.Menu mnu_Import_Data 
      Caption         =   "Import"
   End
   Begin VB.Menu mnu_Import_Excel 
      Caption         =   "Import Excel Files"
   End
   Begin VB.Menu mnu_Pension 
      Caption         =   "Pen&sion"
   End
   Begin VB.Menu mmnu_Windows 
      Caption         =   "&Windows"
      Begin VB.Menu mmnu_Arrange 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mmnu_Cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mmnu_HTile 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mmnu_VTile 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mmnu_Win_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_sysinfo 
         Caption         =   "&System Information"
      End
      Begin VB.Menu mmnu_Win_sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "About &info:HR"
      End
      Begin VB.Menu mmnu_Help 
         Caption         =   "&Help"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExistIHRWFCEXE As Boolean
Dim ExistSN2322EXE As Boolean
Dim FirstTime As Boolean
Dim xIntellisolMatrix As Boolean

Private Type Commands
    Parent As String
    Key As String
    name As String
    IconKey As String
End Type

Dim NbrCmds As Byte
'Dim IHRNodes As New Collection
Dim xORG As Integer

Const mcsWordApplication   As String = "Word.Application"
Dim bGotItWord As Boolean
Dim oWordApplication As Object

Const mcsExcelApplication   As String = "Excel.Application"
Dim bGotItExcel As Boolean
Dim oExcelApplication As Object
Dim oExcelWorkbook As Object
Dim gdbAdoIHRDS As New ADODB.Connection
Dim rsIHRDS As New ADODB.Recordset
Dim xAdoIHRDB

'Ticket #20589 Franks 07/07/2011 - begin
Dim fglbESQLQ, fglbVSQLQ, fglbPosGrp
Dim rsEntMain As New ADODB.Recordset
Dim snapEntitle As New ADODB.Recordset
Dim xmedLTServ(24)
Dim xmedGTServ(24)
Dim xmedPension(24)
'Ticket #20589 Franks 07/07/2011 - end

'Private Const LOCALE_SSHORTDATE            As Long = &H1F    'short date
Private Const LOCALE_SSHORTDATE = &H1F 'short date
Private Declare Function GetLocaleInfo Lib "kernel32.dll" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long

Private Declare Function SetLocaleInfo Lib "kernel32.dll" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


Private Sub mnu_PensionPct_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSPenEnt
        frmSPenEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub cmdForLostFocus_GotFocus()
    If Me.ActiveForm Is Nothing Then
        set_Buttons
        
        '7.9 Picture
        lstPanel.Visible = True
        lstView.Visible = True
        
    End If
End Sub

Private Sub fraMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xORG = X
End Sub

Private Sub fraMove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim xMove
    
    xMove = X - xORG
    If tvwTree.Width + xMove < 0 Then Exit Sub
    If lstView.Width - xMove < 0 Then Exit Sub
    tvwTree.Width = tvwTree.Width + xMove
    panTree.Width = tvwTree.Width + 200
    fraMove.Left = fraMove.Left + xMove
    lstView.Width = lstView.Width - xMove
    lstPanel.Width = lstView.Width
End Sub

Private Sub chkFollow_Ups()
On Error GoTo err_chkFollow_Ups

    panHelp(0) = lStr("Checking for Follow-Up Messages - stand by")
    Screen.MousePointer = HOURGLASS
    
    If gSec_Inq_Follow_Ups And glbFOLLOWUPS Then Load frmvFOLOWUP
    
    If Not glbFollowUpsRemain And Not glbFollwUpsFound Then
        Unload frmvFOLOWUP
        MDIMain.panHelp(0).Caption = "Select an option from menu above"
    Else
        frmvFOLOWUP.Show
    End If

Screen.MousePointer = DEFAULT

Exit Sub
err_chkFollow_Ups:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "chkFollow_Ups", "frmvFOLOWUP")
    Resume Next
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub lstView_GotFocus()
    If Me.ActiveForm Is Nothing Then
        Call set_Buttons
    End If
End Sub

Private Sub MainToolBar_Click()
Dim ctl As Control
On Error Resume Next
    If Not Me.ActiveForm Is Nothing Then
        Me.cmdForLostFocus.SetFocus
        DoEvents
    End If
End Sub

Private Sub MDIForm_Load()
Dim Msg As String
Dim rsEmp As New ADODB.Recordset
Dim xFile, xIHRDS, SQLQ
Dim xEmpCnt As Integer
Dim xAttCnt As Integer

On Error GoTo err_open_manform

'Call SetLocaleInfo(1033, Nothing, Nothing)
If glbFrench = True Then
    Call SetLocaleInfo(1033, LOCALE_SSHORTDATE, "dd/MM/yyyy")
End If

frmFind = False
glbUS = False

If Not glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation - Ticket #10844
    mnu_Import_Excel.Visible = False
End If

If glbCountry = "U.S.A." Then glbUS = True
Me.Height = 8000
Me.Width = 10000
Me.Icon = frmSPLASH.Icon
Me.WindowState = 2

MDIMain.panHelp(0).Caption = "Select an option from menu above"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

'Data Source Name - Begin
If Dir(glbIHRREPORTS & "IHRLin.exe") = "" Then 'Ticket #12564
    xIHRDS = ""
    If Not Dir(glbIHRREPORTS & "IHRDS.mdb") = "" Then
        xAdoIHRDB = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & glbIHRREPORTS & "IHRDS.mdb"
        If gdbAdoIHRDS.State = adStateOpen Then gdbAdoIHRDS.Close
        gdbAdoIHRDS.CommandTimeout = 600
        gdbAdoIHRDS.Mode = adModeReadWrite
        gdbAdoIHRDS.Open xAdoIHRDB
        SQLQ = "SELECT * FROM HR_DATA_SOURCE WHERE DS_SERVER = '" & SQLServerName & "' "
        If Not glbOracle Then
            SQLQ = SQLQ & "AND DS_DATABASE = '" & SQLDatabaseName & "' "
        End If
        If rsIHRDS.State <> 0 Then rsIHRDS.Close
        rsIHRDS.Open SQLQ, gdbAdoIHRDS, adOpenKeyset, adLockOptimistic
        If Not rsIHRDS.EOF Then
            'xIHRDS = " [" & rsIHRDS("DS_NAME") & "] "
            xIHRDS = " " & rsIHRDS("DS_NAME") & " "
        End If
        rsIHRDS.Close
    End If
End If
'Data Source Name - End

'Version = App.Major & "." & App.Minor & "." & App.Revision

If Len(xIHRDS) > 0 Then
    If glbMulti Then
        Me.Caption = "info:HR " & App.Major & "." & App.Minor & "." & App.Revision & " - Multi Position Module" & xIHRDS
    Else
        Me.Caption = "info:HR " & App.Major & "." & App.Minor & "." & App.Revision & " - " & xIHRDS
    End If
Else
    If glbMulti Then Me.Caption = "info:HR " & App.Major & "." & App.Minor & "." & App.Revision & " - Multi Position Module"
End If

glbCompNo = "001" ' this should be set with number of ees allowed
            ' and company name for headings
If setCompInfo(glbCompNo) = False Then
    Load frmComp
End If
If glbAxxent Then
    Call CreateHRRSP
End If
If glbCBrant Then
    Call CreateVacBrant
End If

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
glbUserEmpNo = 0
glbUserEmpNo = Get_UserID_Info(glbUserID, "EMPNBR", 0)


Call setLabels

lblTime = Format(Now, "Short Date") & " - " & Format(Now, "Medium Time")
'mmnu_Opus_Payroll.Enabled = False 'Jaddy 10/28/99
If Not glbSQL And Not glbOracle Then 'Frank 05/08/03
    If Left(glbIHRDBO, 8) <> "00000000" Then 'Jaddy 10/28/99
        If Dir(glbIHRDBO) <> "" Then xIntellisolMatrix = True 'mmnu_Opus_Payroll.Enabled = True        'Jaddy 10/28/99
    End If 'Jaddy 10/28/99
Else
    If ExistTable(gdbAdoIhr001, "paycode_infohr") Then
        'mmnu_Opus_Payroll.Enabled = True
        xIntellisolMatrix = True
    End If
End If

If glbNDepts = 0 Then
    Msg = "You do not have any departmental frmSECURITY."
    Msg = Msg & Chr(10) & "- you can not look up Employee Info."
    MsgBox Msg
End If

If Custom_Feature("Check") Then
    mmnu_Custom.Visible = True
Else
    mmnu_Custom.Visible = False
End If

glbFollowUpsRemain% = False ' on entry don't display followup
                            ' form if no records found
If glbWFC Then
    On Error GoTo EmailSetupError
    'Comment by Frank Ticket# 9690
    'rsEmp.Open "SELECT ED_LOC FROM HREMP WHERE ED_EMPNBR=" & glbEmpNbr, gdbAdoIhr001
    'If Not rsEmp.EOF And Not rsEmp.BOF Then
    '    Select Case UCase(rsEmp("ED_LOC"))
    '        Case "TROY"
    '            glbSMTPServerIP = "10.2.3.15"
    '        Case "KIPL"
    '            glbSMTPServerIP = "10.3.3.2"
    '        Case "ATLA"
    '            glbSMTPServerIP = "10.4.3.2"
    '        Case "TILB"
    '            glbSMTPServerIP = "10.6.3.3"
    '        Case "FREM"
    '            glbSMTPServerIP = "10.7.3.2"
    '        Case "MORV"
    '            glbSMTPServerIP = "10.10.3.1"
    '        Case "WHBY"
    '            glbSMTPServerIP = "10.11.3.2"
    '        Case "BROD"
    '            glbSMTPServerIP = "10.12.3.1"
    '        Case "STJR"
    '            glbSMTPServerIP = "10.13.3.4"
    '        Case "SARN"
    '            glbSMTPServerIP = "10.19.3.3"
    '        Case "ROM"
    '            glbSMTPServerIP = "10.8.3.2"
    '    End Select
    'End If
    '
    'rsEmp.Close
End If

'CHECKED mnu_Pos_Budget.Visible = glbWHSCC
'CHECKED mmnu_R_PlanEstablishment.Visible = glbWHSCC   'Removing Reports from Menu Bar
'mmnu_EE_ASL.Visible = glbWHSCC
'mnuUnionSickBank.Visible = glbWHSCC

ExistIHRWFCEXE = False

If glbWFC Then
    xFile = App.Path
    xFile = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    xFile = xFile & "IHRWFC.exe"
    If Not (Dir(xFile) = "") Then
        ExistIHRWFCEXE = True
    End If
Else
    mnu_Pension.Visible = False
End If

ExistSN2322EXE = False

If glbGuelph Then
    xFile = App.Path
    xFile = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    xFile = xFile & "SN2322.exe"
    If Not (Dir(xFile) = "") Then
        ExistSN2322EXE = True
    End If
End If

AfterEmailSetup:

'Ticket #20209 Franks 04/21/2011, create a new function isIHRPWeb to set glbPayWeb
'' danielk - 01/02/03 - PayWeb interface
'If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRPWeb.EXE") <> "" Then
'    glbPayWeb = True
'    glbPayWebEXE = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRPWeb.EXE" '
'Else
'    glbPayWeb = False
'End If

'If glbWFC Then
'    If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
'        glbIntegrationEXE = "u:\HR Systems VB6\IHRIntegration\IHRIntegrationWFC.EXE" '
'    Else
'        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRIntegrationWFC.EXE") <> "" Then
'            glbIntegrationEXE = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRIntegration.EXE" '
'        End If
'    End If

'Else
    If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
        glbIntegrationEXE = "u:\HR Systems VB6\IHRIntegration\IHRIntegration.EXE" '
    Else
        'Ticket #15223
        If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRIntegration.EXE") <> "" And (Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRIntegration.DAT") <> "" Or Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "MediPay.DAT") <> "" Or Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "etpath.DAT") <> "" Or Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "greatplains.DAT") <> "") Then
            glbIntegrationEXE = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRIntegration.EXE" '
        End If
    End If
'End If
' danielk - 01/02/03 - end

If glbWFC Then
    ''mmnu_Custom.Visible = False
    'Ticket #20305 Franks 05/17/2011 - do not use NGS_Trans.dat, always is on
    'xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "NGS_Trans.dat"
    'If Dir(xFile) <> "" Then
    '    glbNGS_OnFlag = True
    'Else
    '    glbNGS_OnFlag = False
    'End If
    glbNGS_OnFlag = True
    
    ''Ticket #23247 Franks 04/22/2013
    'xFile = glbIHRREPORTS & "WFC_US_Ben_Trans.dat"
    'If Dir(xFile) <> "" Then
    '    glbWFC_US_Ben_Trans = True
    'Else
    '    glbWFC_US_Ben_Trans = False
    'End If
    glbWFC_US_Ben_Trans = True
    
    'If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'Ticket #29013 Franks 08/23/2016 - turn on for programming only, will turn on for WFC once finish all works
    '    glbWFC_IncentivePlanFlag = True
    'End If
    
    'xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "WFC_IncentivePlan.dat"
    'If Dir(xFile) <> "" Then
    '    glbWFC_IncentivePlanFlag = True
    'Else
    '    glbWFC_IncentivePlanFlag = False
    'End If
    glbWFC_IncentivePlanFlag = True
    
End If

If glbCompSerial = "S/N - 2439W" Then 'OK Tire Ticket #21518 Franks 04/26/2012
    'If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
    '    glbIsGWL = True '
    'Else
    '    glbIsGWL = False 'will turn it on for customer when all functions are done
    'End If
    glbIsGWL = True 'Franks 09/20/2012
End If

'Ticket #25208 - New Employee option is available only if you have right to add a New Hire
mnu_NewEmployee.Enabled = gSec_Add_NewHire

'Now Import/Export is for everyone
'If Not glbLinamar Then
    'Ticket #12728
    mmnu_ImportExport.Visible = File_Exist("IHREI.exe")
'End If

If glbSQL Or glbOracle Then
    mmnu_AppT.Visible = File_Exist("IHRAppTrack.exe") And File_Exist("IHRAppTrack.dat")  '7.9 Enhancement - original 'IhrAppT.exe'
Else
    mmnu_AppT.Visible = File_Exist("IHRAPPT.MDB.exe")
End If

If glbLinamar Then
    mnu_F_Division.Visible = True
    mnu_F_Division.Caption = lStr(mnu_F_Division.Caption)
    'MainToolBar.ButtonS(18).ButtonMenus(5).Text = lStr(MainToolBar.ButtonS(18).ButtonMenus(5).Text)
    'MainToolBar.ButtonS(18).ButtonMenus(5).Visible = True
    'Ticket #28118 Franks 02/01/2016
    MainToolBar.ButtonS(18).ButtonMenus(6).Text = lStr(MainToolBar.ButtonS(18).ButtonMenus(5).Text)
    MainToolBar.ButtonS(18).ButtonMenus(6).Visible = True
End If

''have trouble to open it, leave it for now
''If glbWFC Then 'Ticket #25676 Franks 07/29/2014
''    ''mnu_F_Candidate.Visible = True
''    ''mnu_F_Candidate.Caption = lStr(mnu_F_Candidate.Caption)
''    MainToolBar.ButtonS(18).ButtonMenus(6).Text = lStr(MainToolBar.ButtonS(18).ButtonMenus(6).Text)
''    MainToolBar.ButtonS(18).ButtonMenus(6).Visible = True
''End If

If glbWFC Then 'Ticket #28118 Franks 02/01/2016
    mnu_F_JobMaster.Visible = True
    MainToolBar.ButtonS(18).ButtonMenus(4).Visible = True
End If

'Default the Import Date as non visibility
mnu_Import_Data.Visible = IIf(UCase(Left(App.Path, 10)) = "C:\SSWORK\", True, False) Or (glbCompSerial = "S/N - 9999W" And Date <= CVDate(GetMonth("May") & " 10, 2004"))
'mnu_Import_Data.Visible = True

lstPanel.Width = Me.Width
lstView.Width = lstPanel.Width

'Call TreeSetting
Call remNode
Call set_Buttons

'Ticket #16145 - Turn-Off the daily Vacation Update logic
'For Mitchel Plastics - Ultra Manufacturing
'If glbCompSerial = "S/N - 2335W" Then
'    Call Calculate_Vacation_Entitlement
'End If

'Ticket #20589 - Franks 07/07/2011
'For Samuel
If glbCompSerial = "S/N - 2382W" Then
    'If isSamuelPenCal Then 'Ticket #22509 Franks 09/27/2012 - comment out this line
    'Ticket #21160 Frank 11/17/2011, Muhammad need this to this function on Test System
    'Will disable the isSamuelPenCal function when they goes live for Calculate_Entitlement
        Call Calculate_Entitlement
    'End If
End If
'call this function from frmSECURITY


'If Not glbWFC Then
If glbWFC And glbWFCFullRights Then 'Ticket #29720 Franks 01/19/2017
    'don't call this form for MZ and other users with all plants access rights, it took very long time to open it.
Else
    Call chkFollow_Ups
    MDIMain.panHelp(0).Caption = "Select an option from menu above"
End If
'End If

''Ticket #28373 Franks 03/30/2016
''ESS to Tracker
If glbWFC Then
    xEmpCnt = ESS_To_Tracker_EMP
    xAttCnt = ESS_To_Tracker_ATT
End If

Exit Sub

EmailSetupError:
    If rsEmp.State <> adStateClosed Then rsEmp.Close
    Resume AfterEmailSetup
err_open_manform:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "Mainform", "SELECT")
Resume Next
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Ticket #27456 - close opened Word and Excel Applications from info:HR
    If bGotItWord Then oWordApplication.Quit
    If bGotItExcel Then oExcelApplication.Quit
    Set oWordApplication = Nothing
    Set oExcelApplication = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Sub mmnu_Active_Click()
On Error GoTo Err_Active
    Screen.MousePointer = HOURGLASS
    glbtermopen = False
    glbTERM_Seq = 0
    glbLEE_ID = 0
    
    UnloadFrms
    Call remNode
    
    Screen.MousePointer = DEFAULT
    
Exit Sub
Err_Active:
    If Err = 364 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err

    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_AppT_Click()
    Dim xFile, xPath
    xFile = App.Path
    xFile = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    xPath = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    
    xFile = xPath & "IhrAppTrack.exe"
    If Dir(xFile) = "" Then
        MsgBox xFile & " not found"
    Else
        Call Shell(xFile & " " & glbUserID & "," & glbTxtPassword)
    End If

End Sub

Private Sub mmnu_Arrange_Click()
    MDIMain.Arrange 3
End Sub

Private Sub mmnu_Attend_History_Click()
On Error GoTo Err_Att1
    Screen.MousePointer = HOURGLASS
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Attendance_History Then
        If xAttendance <> "Attendance_History" Then Unload frmVATTEND
        xAttendance = "Attendance_History"
        Load frmVATTEND
        frmVATTEND.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
    Screen.MousePointer = DEFAULT
    
Exit Sub
Err_Att1:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_Cascade_Click()
    MDIMain.Arrange 0
End Sub

Private Sub mmnu_Cobra_Click()
On Error GoTo Err_Cobra

    Load frmCobra
    frmCobra.ZOrder 0
    
Exit Sub
Err_Cobra:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_Company_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        Load frmComp
        frmComp.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_CourseCode_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_CourseCodeMaster Then
        'Screen.MousePointer = HOURGLASS
        'Load frmMCourseCode
        'frmMCourseCode.ZOrder 0
        'Screen.MousePointer = DEFAULT
        Call Get_CourseCode(True)
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Company_Preference_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_CompanyPreference Then
        Screen.MousePointer = HOURGLASS
        Load frmComPrefer
        frmComPrefer.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_General_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_CompanyPreference Then
        Screen.MousePointer = HOURGLASS
        Unload frmComPrefer
        Load frmComPrefer
        frmComPrefer.Caption = "Company Preference - General"
        frmComPrefer.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Email_Notification_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_CompanyPreference Then
        Screen.MousePointer = HOURGLASS
        Unload frmComPrefer
        Load frmComPrefer
        frmComPrefer.Caption = "Company Preference - Email Notifications"
        frmComPrefer.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_FileLocation_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_CompanyPreference Then
        Screen.MousePointer = HOURGLASS
        Unload frmComPrefer
        Load frmComPrefer
        frmComPrefer.Caption = "Company Preference - File Locations"
        frmComPrefer.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_BTI_QuarterEnd_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        glbFormCaption = "Quarter End"
        Unload frmUQuarterEnd
        Load frmUQuarterEnd
        frmUQuarterEnd.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_BTI_YTDCarryover_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        glbFormCaption = "Year End Carryover"
        Unload frmUQuarterEnd
        Load frmUQuarterEnd
        frmUQuarterEnd.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_BTI_YTDReduction_BD_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        glbFormCaption = "Year End Reduction For BD"
        Unload frmUYTDBTI
        Load frmUYTDBTI
        frmUYTDBTI.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_BTI_YTDReduction_NonBD_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        glbFormCaption = "Year End Reduction For Non BD"
        Unload frmUQuarterEnd
        Load frmUQuarterEnd
        frmUQuarterEnd.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PerformanceReview_Click()
    If glbSetPer = True Then Unload frmEPERFORMReview
    glbSetPer = False
    Screen.MousePointer = HOURGLASS
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_Performance Then
        Screen.MousePointer = HOURGLASS
        Unload frmEPERFORMReview
        Load frmEPERFORMReview
        frmEPERFORMReview.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PerformanceReviewH_Click()
    If glbSetPer = False Then Unload frmEPERFORMReview
        glbSetPer = True
        Screen.MousePointer = HOURGLASS
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_Performance Then
        Screen.MousePointer = HOURGLASS
        Load frmEPERFORMReview
        frmEPERFORMReview.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_SF_Download_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSFDownload
    frmSFDownload.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_SF_XMLRPT_Click()
    glbFrmCaption$ = "XML Working Table eport"
    Unload frmSFRXMLTable
    Load frmSFRXMLTable
    frmSFRXMLTable.ZOrder 0
End Sub
Private Sub mmnu_SF_FTPSETUP_Click()
    Screen.MousePointer = HOURGLASS
    glbFrmCaption$ = "FTP Setup"
    Unload frmSFSetup
    Load frmSFSetup
    frmSFSetup.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_SF_XML_LOCATION_Click()
    Screen.MousePointer = HOURGLASS
    glbFrmCaption$ = "XML File Location Setup"
    Unload frmSFSetup
    Load frmSFSetup
    frmSFSetup.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

'Ticket #26912 Franks 06/22/2015 - begin
Private Sub mmnu_Sys247_Download_Click()
    Screen.MousePointer = HOURGLASS
    glbFrmCaption$ = "Download File from FTP"
    Unload frmSFDownload
    Load frmSFDownload
    frmSFDownload.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_Sys247_Upload_Click()
    Screen.MousePointer = HOURGLASS
    glbFrmCaption$ = "Upload File To FTP"
    Unload frmSFDownload
    Load frmSFDownload
    frmSFDownload.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_Sys247_FTPSETUP_Click()
    Screen.MousePointer = HOURGLASS
    glbFrmCaption$ = "FTP Setup"
    Unload frmSFSetup
    Load frmSFSetup
    frmSFSetup.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub
'Ticket #26912 Franks 06/22/2015 - end

Private Sub mmnu_EmlSetup_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_Other_Entitlements Then
        Screen.MousePointer = HOURGLASS
        Load frmUEmergLeave
        frmUEmergLeave.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Counsel_Absence_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        glbFormCaption = "Absence Counseling Setup"
        Unload frmMCounselSet
        Load frmMCounselSet
        frmMCounselSet.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Counsel_LE_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        glbFormCaption = "L/LE Counseling Setup"
        Unload frmMCounselSet
        Load frmMCounselSet
        frmMCounselSet.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Custom_Click()
    Call Custom_Feature("Run")
End Sub

Private Sub mnu_DailyManpower_Click()
    On Error GoTo Eh
    'added by Bryan 9/Sep/05 Ticket #9235

    Screen.MousePointer = HOURGLASS
    
    Load frmRDailyMan
    frmRDailyMan.ZOrder 0

    Screen.MousePointer = DEFAULT
        
Exit Sub
Eh:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mnu_GPPosting_Click()
    On Error GoTo Eh
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        Load frmRGPPosting
        frmRGPPosting.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

    Screen.MousePointer = DEFAULT
    
Exit Sub
Eh:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Function File_Exist(FileName As String) As Boolean
    Dim xFile, xPath
    xFile = App.Path
    xFile = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    xPath = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    
    xFile = xPath & FileName '"IhrAppT.exe"
    If Dir(xFile) = "" Then
        File_Exist = False
    Else
        File_Exist = True
    End If
End Function

Private Function Custom_Feature(CR As String)
'CR, Check or Run
Dim xFile, xPath
xPath = App.Path
xPath = xPath & IIf(Right(xPath, 1) = "\", "", "\")
xFile = "Custom Feature"
Custom_Feature = False

If glbLinamar Then
    xFile = xPath & "IHRLin.exe"
End If
If glbWFC Then
    xFile = xPath & "IHRWFC.exe"
End If
If glbSamuel Then 'Ticket #20885 Franks 12/12/2011
    If gSec_SAM_Show_CustomFeatures Then 'Ticket #21496 Franks 01/26/2012
        xFile = xPath & "IHRSamuel.exe"
    End If
End If
If glbGuelph Then
    xFile = xPath & "SN2322.exe"
End If
If glbCompSerial = "S/N - 2214W" Then 'Casey House
    xFile = xPath & "HRCASEY.exe"
End If
If glbCompSerial = "S/N - 2170W" Then 'London Health Unit
    xFile = xPath & "IHR01.exe"
End If
If glbCompSerial = "S/N - 2217W" Then 'City of Pickering
    xFile = xPath & "IHRPK01.exe"
End If
If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Fulls
    xFile = xPath & "SN2276.exe"
End If
If glbCompSerial = "S/N - 2292W" Then 'City of Elgin
    xFile = xPath & "SN2292.exe"
End If
If glbCompSerial = "S/N - 2242W" Then 'CCAC London
    xFile = xPath & "SN2242.exe"
End If
If glbCompSerial = "S/N - 2339W" Then 'NTN Bearing
    xFile = xPath & "SN2339.exe"
End If
If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab - Ticket #15776
    xFile = xPath & "SN2394.exe"
End If
If glbCompSerial = "S/N - 2349W" Then 'CGL Manufacturing
    If CR = "Run" Then
        frmODBCLogon.CGLInterface = True
        frmODBCLogon.Show
    End If
    
    Custom_Feature = True
    Exit Function
End If
If glbCompSerial = "S/N - 2355W" Then 'County of Lambton
    xFile = xPath & "SN2355.exe"
End If

If Dir(xFile) = "" Then
    If CR = "Run" Then MsgBox xFile & " not found"
Else
    If glbCompSerial = "S/N - 2214W" Or glbCompSerial = "S/N - 2394W" Or glbSamuel Then 'Casey House
        If CR = "Run" Then Call Shell(xFile & " " & glbUserID & "," & DecryptPassword(glbPassword))
    Else
        If CR = "Run" Then Call Shell(xFile & " " & glbUserID & "," & glbPassword)
    End If
    Custom_Feature = True
End If
End Function

Private Sub mmnu_CustomReport_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_CustomReport Then
        frmSCusRPTs.Show
        frmSCusRPTs.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Department_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Departments Then
        Screen.MousePointer = HOURGLASS
        Call Get_Dept(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_OHRSDepartment_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Departments Then
        Screen.MousePointer = HOURGLASS
        Call Get_OHRSDept(True) ' master call of OHRS departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_BonusDepartment_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Departments Then
        Screen.MousePointer = HOURGLASS
        Call Get_DeptBonus(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Disciplinary_Steps_Click()
    Dim SQLQ As String

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If Not gSec_Inq_Master_Table("CERE") Then
    '    MsgBox "You Do Not Have Authority For This Transaction"
    '    Exit Sub
    'End If
    frmMDiscipSteps.Show 1
End Sub

Private Sub mnu_FindJob_Click() 'Ticket #25911 Franks 09/24/2014.
        Screen.MousePointer = HOURGLASS
        'Call Get_Div(True) ' master call of Job
        Call Get_JobMaster(False) ' master call of Job
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Division_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Divisions Then
        Screen.MousePointer = HOURGLASS
        Call Get_Div(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_SalDist_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_SalDist Then
        Screen.MousePointer = HOURGLASS
        Call Get_SalDist(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayCategory_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Payroll_Category Then
        Screen.MousePointer = HOURGLASS
        Call Get_PayCategory(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_ChargeCode_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Charge_Code Then
        Screen.MousePointer = HOURGLASS
        Call Get_ChargeCode(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_ProjectCode_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Project_Code Then
        Screen.MousePointer = HOURGLASS
        Call Get_ProjectCode(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Machine_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Machine Then
        Screen.MousePointer = HOURGLASS
        Call Get_Machine(True) ' master call of departments
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
      
Private Sub mmnu_Multiple_Data_Source_Click()
Dim SQLQ As String, xFile As String
On Error GoTo GPayroll_Err

If glbIsUseIHRDS Then 'Ticket #20310 Franks 05/10/2011
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_MultiDataSourceSetup Then
        xFile = glbIHRREPORTS & "IHRDS.mdb"
        If Dir(xFile) = "" Then
            MsgBox xFile & " not found."
        Else
            Screen.MousePointer = HOURGLASS
            Load frmMDataSource
            frmMDataSource.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End If
Screen.MousePointer = DEFAULT

Exit Sub
GPayroll_Err:
    glbFrmCaption$ = "Get Payroll"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Multiple Data Sources", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_EE_ASL_Click()
On Error GoTo Err_Att
Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_WHSCC_ASL Then
        Load frmVASL
        frmVASL.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Att:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Associations_Click()
On Error GoTo Err_Asso

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Associations Then
        Load frmEASSOC
        frmEASSOC.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Asso:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Training_List_Click()
On Error GoTo Err_Asso

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_UserDefineTbl Then
        Load frmETRAINLST
        frmETRAINLST.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
    
Screen.MousePointer = DEFAULT

Exit Sub
Err_Asso:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_UserDefine_Table_Click()
On Error GoTo Err_Asso

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_UserDefineTbl Then
        Load frmEUserDef
        frmEUserDef.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
    
Screen.MousePointer = DEFAULT

Exit Sub
Err_Asso:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_ADP_Click()
On Error GoTo Err_Att

Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_ADP_Data Then
        Load frmEmpADP
        frmEmpADP.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Att:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Add_PayrollID_Data_Click()
On Error GoTo Err_Att

Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_AddPayrollIDData Then
        Load frmEAddPayrollIDData
        frmEAddPayrollIDData.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Att:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Flags_Click()
On Error GoTo Err_Att

Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_EMP_FLAGS Then
        Load frmEmployeeFlags
        frmEmployeeFlags.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Att:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Attendance_Click()
On Error GoTo Err_Att

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Attendance Then
        If xAttendance <> "Attendance" Then Unload frmVATTEND
        xAttendance = "Attendance"
        
        Load frmVATTEND
        frmVATTEND.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Att:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_WorkSchedule_Click()
On Error GoTo Err_WorkSchedule
    Screen.MousePointer = HOURGLASS
    Load frmEScheduler
    frmEScheduler.ZOrder 0
    Screen.MousePointer = DEFAULT

Exit Sub
Err_WorkSchedule:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    'If Err = 364 Then Exit Sub
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_EmpOther_Click()
On Error GoTo Err_Basic
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_OtherInformation Then
        Screen.MousePointer = HOURGLASS
        Load frmEmpOther
        frmEmpOther.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Basic:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    'If Err = 364 Then Exit Sub
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_EmpHistory_Click()
On Error GoTo Err_Basic
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_EMP_HISTORY Then
        Screen.MousePointer = HOURGLASS
        Load frmEHistory
        frmEHistory.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Basic:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    'If Err = 364 Then Exit Sub
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Languages_Click()
On Error GoTo Err_Basic
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_EMP_LANG Then
        Screen.MousePointer = HOURGLASS
        Load frmELang
        frmELang.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Basic:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    'If Err = 364 Then Exit Sub
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Succession_Click()
On Error GoTo Err_Basic
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_SUCCESSION Then
        Screen.MousePointer = HOURGLASS
        Load frmESuccession
        frmESuccession.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Basic:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    'If Err = 364 Then Exit Sub
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Basic_Click()
On Error GoTo Err_Basic
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Basic Then
        Screen.MousePointer = HOURGLASS
        Load frmEEBASIC
        frmEEBASIC.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Basic:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    'If Err = 364 Then Exit Sub
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Benefits_Click()
On Error GoTo Err_Bene

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Benefits Then
         Unload frmEBENEFITS
         frmEBENEFITS.Show
         frmEBENEFITS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Bene:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Cobra_Click()
On Error GoTo Err_Comm

Screen.MousePointer = HOURGLASS
    
    Load frmCobra
    frmCobra.ZOrder 0

Screen.MousePointer = DEFAULT

Exit Sub
Err_Comm:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Comments_Click()
On Error GoTo Err_Comm

Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Comments Then
    
        Load frmECOMMENTS
        frmECOMMENTS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Screen.MousePointer = DEFAULT

Exit Sub
Err_Comm:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_SalDist_Click()
On Error GoTo Err_Comm

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_SalDist Then
        Load frmESalDist
        frmESalDist.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Comm:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_comPlan_Click()
On Error GoTo Err_Sala

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        Load frmEComPlan
        frmEComPlan.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Sala:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mnuManpower_Click()
'added by Bryan 12/07/05 #8922
    Screen.MousePointer = HOURGLASS
    
        Load frmManpower
        frmManpower.ZOrder 0
'    Else
'        MsgBox "You Do Not Have Authority For This Transaction"
'    End If
    Screen.MousePointer = DEFAULT
    
    Exit Sub
    
Err_Sala:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mnu_Manpower_Plan_Click()
    'added by Bryan 14/07/05 #8922
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Manpower_Plan Then
        Screen.MousePointer = HOURGLASS
        
        Load frmRManpower
        frmRManpower.ZOrder 0
    
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You do not have Authority for this transaction"
    'End If
Exit Sub
Err_Sala:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Contact_Click()
On Error GoTo Err_Contact

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Basic Then
        Load frmEMERG
        frmEMERG.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Exit Sub
Err_Contact:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Counsel_Click()
    On Error GoTo Err_Cou
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Counselling Then
        Load frmECounsel
        frmECounsel.ZOrder 0
    'Else
    '    MsgBox "You do not have authority for this transaction."
    'End If
Exit Sub
    
Err_Cou:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Dependants_Click()
On Error GoTo Err_Dep
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Dependents Then
        Load frmDEPNDTS
        frmDEPNDTS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Exit Sub
Err_Dep:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_EFollowup_Click()
On Error GoTo Err_Fol1

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Follow_Ups Then
        Load frmEFOLLOWUP
        frmEFOLLOWUP.ZOrder 0
        'frmEFOLLOWUP.WindowState = 2
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Fol1:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_FormalEd_Click()

Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Formal_Education Then
        Load frmFORMALED
        frmFORMALED.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_EE_HS_Acci_Cost_Click()
On Error GoTo Err_Status

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        Unload frmEHSCost
        Load frmEHSCost
        frmEHSCost.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Status:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_HS_Company_Costs_Click()
Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        Unload frmEHSCOMPCost
        Load frmEHSCOMPCost
        frmEHSCOMPCost.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_EE_HSDiv_Contact_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = True
        Unload frmEHSContact
        Load frmEHSContact
        frmEHSContact.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Contact_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSContact
        Load frmEHSContact
        frmEHSContact.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HSDiv_Corrective_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = True
        Unload frmEHSCorrective
        Load frmEHSCorrective
        frmEHSCorrective.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Corrective_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSCorrective
        Load frmEHSCorrective
        frmEHSCorrective.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Cost_Click()
On Error GoTo Err_HS4

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        Unload frmEHSWCBC
        Load frmEHSWCBC
        frmEHSWCBC.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS4:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HSDiv_Incident_Data_Click()
On Error GoTo Err_HS1

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = True
        Unload frmEHSINCIDENT
        Load frmEHSINCIDENT
        frmEHSINCIDENT.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS1:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_W7_CompanyMaster_Click()
On Error GoTo Err_HS1a

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmMW7CmpMst
        Load frmMW7CmpMst
        frmMW7CmpMst.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS1a:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Incident_Data_Click()
On Error GoTo Err_HS1

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSINCIDENT
        Load frmEHSINCIDENT
        frmEHSINCIDENT.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS1:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HSDiv_Incident_Documents_Click()
On Error GoTo Err_HS1

Screen.MousePointer = HOURGLASS
    If gsAttachment_DB Then
        'Ticket #16189 - Commented out previous logic for screen load security since now
        'only screens the user has access to will be visible.
        'If gSec_Inq_Health_Safety Then
            glbLinHS = True
            Unload frmEHSAttach 'frmEHSINCIDENTDocuments
            Load frmEHSAttach
            frmEHSAttach.ZOrder 0
        'Else
        '    MsgBox "You Do Not Have Authority For This Transaction"
        'End If
    Else
        MsgBox "Attachment is not setup on Company Preference screen."
    End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS1:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

'George on Feb 3,2005 for Incident Documents attachment
Private Sub mmnu_EE_HS_Incident_Documents_Click()
On Error GoTo Err_HS1

Screen.MousePointer = HOURGLASS
    If gsAttachment_DB Then
        'Ticket #16189 - Commented out previous logic for screen load security since now
        'only screens the user has access to will be visible.
        'If gSec_Inq_Health_Safety Then
            glbLinHS = False
            Unload frmEHSAttach 'frmEHSINCIDENTDocuments
            Load frmEHSAttach
            frmEHSAttach.ZOrder 0
        'Else
        '    MsgBox "You Do Not Have Authority For This Transaction"
       ' End If
    Else
        MsgBox "Attachment is not setup on Company Preference screen."
    End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS1:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Reoccurrence_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        Unload frmEHSReOccur
        Load frmEHSReOccur
        frmEHSReOccur.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HSDiv_Injury_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = True
        Unload frmEHSINJURY
        Load frmEHSINJURY
        frmEHSINJURY.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Injury_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSINJURY
        Load frmEHSINJURY
        frmEHSINJURY.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_InjuryWF7_Click()
On Error GoTo Err_HSWF7

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSINJURYWF7
        Load frmEHSINJURYWF7
        frmEHSINJURYWF7.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HSWF7:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_WSIBF9_Click()
On Error GoTo Err_HSWF9

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSF9
        Load frmEHSF9
        frmEHSF9.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HSWF9:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HSDiv_Medical_Click()
On Error GoTo Err_HS3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = True
        Unload frmEHSWCB
        Load frmEHSWCB
        frmEHSWCB.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Medical_Click()
On Error GoTo Err_HS3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        
        If Not glbLinamar Then
            Unload frmEHSWCB
            Load frmEHSWCB
            frmEHSWCB.ZOrder 0
        Else
            Unload frmEHSEMPWCB
            Load frmEHSEMPWCB
            frmEHSEMPWCB.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HSDiv_Root_Cause_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = True
        Unload frmEHSCause
        Load frmEHSCause
        frmEHSCause.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_HS_Root_Cause_Click()
On Error GoTo Err_HS2

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        glbLinHS = False
        Unload frmEHSCause
        Load frmEHSCause
        frmEHSCause.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_HS2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_ProfitSharing_Click()
On Error GoTo Err_ODoll

    Screen.MousePointer = HOURGLASS

    Unload frmEProfitSharing
    frmEProfitSharing.Show
    frmEProfitSharing.ZOrder 0

    Screen.MousePointer = DEFAULT

Exit Sub
Err_ODoll:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_ODollar_Click()
On Error GoTo Err_ODoll

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Other_Entitlements Then
        Unload frmEODOLLAR
        frmEODOLLAR.Show
        frmEODOLLAR.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_ODoll:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_PayTrans_Click()
On Error GoTo Err_Other

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayrollTrans Then
        Unload frmEPayTran
        frmEPayTran.Show
        frmEPayTran.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Other:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Other_Click()
On Error GoTo Err_Other

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Earnings Then
        Unload frmOTHERERN
        frmOTHERERN.Show
        frmOTHERERN.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Other:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Payroll_Click()
On Error GoTo Err_Bank

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Banking Then
        Screen.MousePointer = HOURGLASS
        
        Load frmEBANK
        frmEBANK.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Bank:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Performance_Click()
On Error GoTo Err_Perf

If glbSetPer = True Then Unload frmEPERFORM
glbSetPer = False

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Performance Then 'changed by RAUBREY 5/15/97
    '    If Not isUpdated(frmEPERFORM) Then Exit Sub
        Unload frmEPERFORM
        frmEPERFORM.Show
        frmEPERFORM.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Perf:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Position_Click()
On Error GoTo Err_Pos

    If glbSetPos = True Then Unload frmEPOSITION
    glbSetPos = False
    
Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Position Then
        Unload frmEPOSITION
        frmEPOSITION.Show
        frmEPOSITION.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Pos:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Temp_CrossTrain_Position_Click()
On Error GoTo Err_TempPos

    'Ticket #16189-------------------------------
    'If glbSetPos = True Then Unload frmETmpCrsTrnPos
    'glbSetPos = False
    'Ticket #16189-------------------------------
    
Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Position Then
        Unload frmETmpCrsTrnPos
        frmETmpCrsTrnPos.Show
        frmETmpCrsTrnPos.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_TempPos:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Salary_Click()
On Error GoTo Err_Sala
    
    If glbCompSerial = "S/N - 2288W" Then 'tkt#10845
        If glbSetSal = True Then Unload frmESALARYMusashi
    ElseIf glbCompSerial = "S/N - 2494W" Then   'Ticket #30452 - ONE CARE
        If glbSetSal = True Then Unload frmESALARY2
    Else
        If glbSetSal = True Then Unload frmESALARY
    End If
    glbSetSal = False
    
Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        'jaddy changed by jerry request
        'If Not isUpdated(frmESALARY) Then Exit Sub
        If glbCompSerial = "S/N - 2288W" Then 'tkt#10845
            Unload frmESALARYMusashi
            frmESALARYMusashi.Show
            frmESALARYMusashi.ZOrder 0
        ElseIf glbCompSerial = "S/N - 2494W" Then   'Ticket #30452 - ONE CARE
            Unload frmESALARY2
            frmESALARY2.Show
            frmESALARY2.ZOrder 0
        Else
            Unload frmESALARY
            frmESALARY.Show
            frmESALARY.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Sala:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_Seminars_Click()
On Error GoTo Err_Sem

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Education_Seminars Then
        Load frmESEMINARS
        frmESEMINARS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Sem:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Skills_Click()
On Error GoTo Err_Skills

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Skills Then
        Load frmESkills
        frmESkills.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Skills:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Skills_Production_Click()
On Error GoTo Err_Skills

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_LinamarSkills Then
        Load frmLinamarSkills
        frmLinamarSkills.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Skills:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_Status_Click()
On Error GoTo Err_Status

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Basic Then
        Load frmEESTATS
        frmEESTATS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Status:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_EE_vWorkFlow_Click()
On Error GoTo Err_Fol2

Screen.MousePointer = HOURGLASS
        Load frmEWorkFlow
        frmEWorkFlow.ZOrder 0

Screen.MousePointer = DEFAULT

Exit Sub
Err_Fol2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_vFollowup_Click()
On Error GoTo Err_Fol2

Screen.MousePointer = HOURGLASS
    
    glbFollowUpsRemain% = True
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Follow_Ups Then
        Load frmvFOLOWUP
        frmvFOLOWUP.ZOrder 0
        'frmVFOLOWUP.WindowState = 2
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Fol2:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_EE_VSE_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Entitlements Then
        ' EEBasic.Show
        Load frmVACSICK
        frmVACSICK.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_EE_VSO_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Entitlements Then
        Load frmVACSICKO
        frmVACSICKO.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_EE_OvertimeO_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Ovt_Overview Then
        Load frmOvtBankO
        frmOvtBankO.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Employees_Click()
    MDIMain.panHelp(0).Caption = "Select a menu item"
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(3).Caption = ""
End Sub

Private Sub mmnu_ImportExport_Click()
Dim xFile, xPath
    xFile = App.Path
    xFile = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    xPath = xFile & IIf(Right(xFile, 1) = "\", "", "\")
    
    xFile = xPath & "IHREI.exe"
    If Dir(xFile) = "" Then
        MsgBox xFile & " not found"
    Else
        'Call Shell(xFile & " " & glbUserID & "," & glbPassword)
        Call Shell(xFile & " " & glbUserID & "," & glbTxtPassword)
    End If
End Sub
' Visible menu bar:

Private Sub mmnu_mail_Click()
    Call mnu_Mail_Click
End Sub

Private Sub mnu_F_Candidate_Click()
    'frmSFFind.Show
    'lstPanel.Visible = True
    'lstView.Visible = True
    frmJOBS.Show '1
    'Call mmnu_FindCandi_Click
End Sub

Private Sub mnu_F_Close_Click()
    If Not Me.ActiveForm Is Nothing Then
        glbUserUploadMode = UploadFormWithCheck: Unload Me.ActiveForm
    End If
    DoEvents
    If Me.ActiveForm Is Nothing Then
        Call set_Buttons
        
        '7.9 - Picture
        lstPanel.Visible = True
        lstView.Visible = True
    End If
End Sub

Private Sub mnu_F_Division_Click()
    Call GET_Division
    lstPanel.Visible = False
    lstView.Visible = False
End Sub

Private Sub mnu_F_Employee_Click()
    Call GET_EMP
    lstPanel.Visible = False
    lstView.Visible = False
End Sub

Private Sub mnu_F_JobMaster_Click()
    Call clkFind(RelateJobMaster)
End Sub
Private Sub mnu_F_Position_Click()
    Call clkFind(RelatePos)
End Sub

Private Sub mnu_F_Preview_Click()
    Call Me.ActiveForm.cmdView_Click
End Sub

Private Sub mnu_F_Print_Click()
    Call Me.ActiveForm.cmdPrint_Click
End Sub

Private Sub mnu_F_Exit_Click()
    Dim xForm As Form
    If Not Me.ActiveForm Is Nothing Then
        Call isUpdated(Me.ActiveForm)
    End If
    DoEvents
    Call ApplicationEnd
End Sub

Private Sub mnu_E_NewRecord_Click()
    Call clkNew("NewRecord")
End Sub

Private Sub mnu_E_Save_Click()
    If Me.ActiveForm Is Nothing Then Exit Sub
    Call Me.ActiveForm.cmdOK_Click
End Sub

Private Sub mnu_E_Cancel_Click()
    If Me.ActiveForm Is Nothing Then Exit Sub
    Call Me.ActiveForm.cmdCancel_Click
End Sub

Private Sub mnu_E_Delete_Click()
    If Me.ActiveForm Is Nothing Then Exit Sub
    Call Me.ActiveForm.cmdDelete_Click
End Sub

Private Sub mnu_F_TEmployee_Click()
    Call mnu_Term_Inquiry_Click
    lstPanel.Visible = False
    lstView.Visible = False
End Sub

Private Sub mnu_Import_Data_Click()
    'frmImportDbSQL.Show
    frmImportDb.Show
End Sub

Private Sub mnu_Import_Excel_Click()
    frmImport.Show
End Sub

Private Sub mnu_M_Add_Click()
    Call Me.ActiveForm.cmdNew_Click
End Sub

Private Sub mnu_WORD_Click()
    bGotItWord = DoLaunchWord(oWordApplication)
End Sub

Private Sub mnu_EXCEL_Click()
    bGotItExcel = DoLaunchExcel(oExcelApplication)
End Sub

Private Sub mnu_HRsoft_Click() 'Ticket #24184 Franks 12/04/2013
Dim theWebSite As String
    'theWebSite = "https://recruitment.workstreaminc.com" ' "http://www.infohr.net"
    'Ticket #27275 Franks 07/06/2015
    theWebSite = "https://recruitment.hrsoft.com"
    If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
        Call Shell("explorer.exe " & theWebSite, vbNormalFocus)
    Else
        Call Shell("W:\Program Files\Google\Chrome\Application\chrome.exe " & theWebSite, vbNormalFocus)
    End If
    
End Sub

Private Sub mnu_NewContractEmployee_Click()
    
    'Ticket #29965 Franks 03/30/2017 - MZ asked to remove the password
    '''Ticket #29828 Franks 02/10/2017 - begin
    '''Jerry: In main menu bar, add a password to the Contract Employee icon. - petman
    ''glbAccessPswd = False
    ''frmAccessPswd.Caption = "Password Required"
    ''frmAccessPswd.lblPassword.Caption = "Enter the Password to add a new Contractor"
    ''frmAccessPswd.Show 1
    ''If glbAccessPswd = False Then   'Access Denied
    ''    Exit Sub
    ''End If
    '''Ticket #29828 Franks 02/10/2017 - end
    
    If glbtermopen Then
        Call mmnu_Active_Click
    End If
    Call clkNew("ContractEmployee")
End Sub

Private Sub mnu_Mail_Click()
    If Not UserEmailExist Then
        Exit Sub
    End If
    If (glbOnTop = "FRMEESTATS" Or glbOnTop = "FRMEMERG" Or glbOnTop = "FRMETERM") And Not (MDIMain.ActiveForm Is Nothing) Then
        Call Me.ActiveForm.imgEmail_Click
    Else
        Call GeneralEmailbox
    End If
End Sub

Private Sub GeneralEmailbox()
Dim SQLQ, xEmail

    xEmail = GetCurEmpEmail
    If Len(xEmail) > 0 Then
        frmSendEmail.txtTo.Text = xEmail
        frmSendEmail.Tag = ""
        frmSendEmail.Show 1
    Else
        If Len(glbLEE_SName) = 0 Then
            MsgBox "There is no email on Status/Dates screen for employee. "
        Else
            MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
        End If
    End If
End Sub

Private Sub mnu_M_Delete_Click()
    Call Me.ActiveForm.cmdDelete_Click
End Sub

Private Sub mnu_M_Down_Click()
    If Not glbSkip Then
        If Me.ActiveForm.name = "frmEESTATS" Then
            'Ticket #18123
            'This has been added because clients are getting prompt to Save Status/Dates screen when they
            'have not made any changes. The reason for this missing corresponding HREMP_OTHER record. By
            'refreshing the data control seems to resolve the issue. In future we can remove this as currently
            'the HREMP_OTHER table is new and not all older employees have corresponding record in this.
            Me.ActiveForm.DataOther.Refresh
        End If
        If Not isUpdated(Me.ActiveForm) Then Exit Sub
    End If
    Call RowDown
End Sub

Private Sub mnu_M_Up_Click()
    If Not glbSkip Then
        If Me.ActiveForm.name = "frmEESTATS" Then
            'Ticket #18123
            'This has been added because clients are getting prompt to Save Status/Dates screen when they
            'have not made any changes. The reason for this missing corresponding HREMP_OTHER record. By
            'refreshing the data control seems to resolve the issue. In future we can remove this as currently
            'the HREMP_OTHER table is new and not all older employees have corresponding record in this.
            Me.ActiveForm.DataOther.Refresh
        End If
        'Me.MainToolBar.ButtonS(16).Enabled = False
        If Not isUpdated(Me.ActiveForm) Then Exit Sub
    End If
    Call RowUp
End Sub

Private Sub mnu_mailbox_Click()
    'Call Me.ActiveForm.cmdDelete_Click
End Sub

Private Sub mmnu_File_Audit_Click()
On Error GoTo FileAuditErr
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Audit Then
        Screen.MousePointer = HOURGLASS
        Me.MainToolBar.ButtonS(10).Visible = True
        Me.MainToolBar.ButtonS(10).Enabled = True
        Load frmAUDIT
        frmAUDIT.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Exit Sub

FileAuditErr:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_Attendance_Audit_Click()
On Error GoTo AttAuditErr
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Audit Then
        Screen.MousePointer = HOURGLASS
        'Me.MainToolBar.ButtonS(10).Visible = False
        'Me.MainToolBar.ButtonS(10).Enabled = True
        If glbCompSerial = "S/N - 2241W" Then ' Granite Club Ticket #16017
            Load frmAttAUDIT
            frmAttAUDIT.ZOrder 0
        Else
            'Release 8.0 - Ticket #22682
            Load frmAUDITAttend
            frmAUDITAttend.ZOrder 0
        End If
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Exit Sub

AttAuditErr:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_File_CounselAudit_Click()
On Error GoTo FileCounselAuditErr
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_CounselAudit Then
        Screen.MousePointer = HOURGLASS
        Me.MainToolBar.ButtonS(10).Visible = True
        Me.MainToolBar.ButtonS(10).Enabled = True
        Load frmAUDITCounsel
        frmAUDITCounsel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Exit Sub

FileCounselAuditErr:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub GroupBenefits()
On Error GoTo Err_GB3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_BenefitGroupSetup Then
        If glbLinamar Then
            MsgBox "This function is not available."
        Else
            Load frmBENGR
            frmBENGR.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_GB3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub GroupManulifeRule()
Dim SQLQ As String
On Error GoTo Err_GB3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_BenefitGroupSetup Then
        If glbLinamar Then
            MsgBox "This function is not available."
        Else
            SQLQ = "SELECT * FROM HR_MANULIFE_TRAN_RULE "
            SQLQ = SQLQ & " ORDER BY MT_SECTION,MT_EMP "
    
            frmSManulifeRule.data1.ConnectionString = glbAdoIHRDB
            frmSManulifeRule.data1.RecordSource = SQLQ
            frmSManulifeRule.data1.Refresh
            
            frmSManulifeRule.Show 1
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_GB3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub GroupOMERS_Formula()
Dim SQLQ As String
Dim rsTest As New ADODB.Recordset
On Error GoTo Err_GB3

Screen.MousePointer = HOURGLASS
'SQLQ = "SELECT * FROM HR_OMERS_FORMULA "
'SQLQ = SQLQ & " ORDER BY OM_YEAR DESC "
'
'frmOmersFormula.Data1.ConnectionString = glbAdoIHRDB
'frmOmersFormula.Data1.RecordSource = SQLQ
'frmOmersFormula.Data1.Refresh
If isTableInDB("HR_OMERS_FORMULA") Then
    frmOmersFormula.Show 1
Else
    MsgBox "Missing table " & "HR_OMERS_FORMULA" & Chr(10) & "Please contact info:HR support."
End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_GB3:
    If Err = 364 Then
        MsgBox "Missing HR_OMERS_FORMULA Table."
        Exit Sub
    End If
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub GroupBenefitMatrix()
Dim SQLQ As String
On Error GoTo Err_GB3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_BenefitGroupSetup Then
        If glbLinamar Then
            MsgBox "This function is not available."
        Else
            SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX "
            SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP,BM_DIV "
    
            frmBenGrpMatrix.data1.ConnectionString = glbAdoIHRDB
            frmBenGrpMatrix.data1.RecordSource = SQLQ
            frmBenGrpMatrix.data1.Refresh
            
            frmBenGrpMatrix.Show 1
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_GB3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub BenefitCost()
On Error GoTo Err_GB3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_BenefitGroupSetup Then
        If glbLinamar Then
            MsgBox "This function is not available."
        Else
            Load frmSBenCost
            frmSBenCost.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_GB3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_BenefitRates()
On Error GoTo Err_GB3

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_BenefitGroupSetup Then
        If glbLinamar Then
            MsgBox "This function is not available."
        Else
            Load frmSBenRates
            frmSBenRates.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_GB3:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_File_Click()
    MDIMain.panHelp(0).Caption = "Select a menu item"
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(3).Caption = ""
End Sub

Private Sub mmnu_File_HrlyEntitlements_Click()
On Error GoTo Err_Hrly
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Hrly_Entitlements Then
        Screen.MousePointer = HOURGLASS
        
        Load frmHrEnt
        frmHrEnt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Hrly:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_File_Pay_Pariod_Master_Click()
Dim SQLQ As String
On Error GoTo GPayroll_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayPeriod_Master Then
        Screen.MousePointer = HOURGLASS
        Load frmPayPeriodMaster
        'frmPayPeriodMaster.Show
        frmPayPeriodMaster.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
GPayroll_Err:
    glbFrmCaption$ = "Get Payroll"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Payroll", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_Close_Pay_Period_Click()
Dim SQLQ As String
On Error GoTo ClosePP_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayPeriod_Master Then
        Screen.MousePointer = HOURGLASS
        Load frmClosePP
        'frmPayPeriodMaster.Show
        frmClosePP.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
ClosePP_Err:
    glbFrmCaption$ = "Get Payroll"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Close Pay Period", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_OvertimeMaster_Click()
Dim SQLQ As String
On Error GoTo OvertimeMaster_Click_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("OvertimeMaster_MassUpdate", glbUserID) Then
        If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
            frmSOvertimeMst.Caption = "Extra Time Master"
        End If
        Screen.MousePointer = HOURGLASS
        Load frmSOvertimeMst
        frmSOvertimeMst.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
OvertimeMaster_Click_Err:
    If glbCompSerial = "S/N - 2425W" Then 'Four Villages (Ticket #19998)
        glbFrmCaption$ = "Extra Time Master"
    Else
        glbFrmCaption$ = "Overtime Master"
    End If
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "OvertimeMst", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_SalaryIncr_Click()
Dim SQLQ As String
On Error GoTo mmnu_SalaryIncr_Click_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        Screen.MousePointer = HOURGLASS
        Load frmSalPerctg
        frmSalPerctg.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
mmnu_SalaryIncr_Click_Err:
    glbFrmCaption$ = "Salary Increase Rule"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "SalaryIncr", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_SeniorityDateCalculation_Click()
Dim SQLQ As String
On Error GoTo mmnu_SeniorityDateCalculation_Click_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        'Screen.MousePointer = HOURGLASS
        Call SeniorityDateCalculation
        'Load frmSenDateCalc
        'frmSenDateCalc.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
mmnu_SeniorityDateCalculation_Click_Err:
    glbFrmCaption$ = "Seniority Date Calculation"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "SenDateCalc", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_VacationIncr_Click()
Dim SQLQ As String
On Error GoTo mmnu_VacationIncr_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        Screen.MousePointer = HOURGLASS
        Load frmVacPerctg
        frmVacPerctg.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
mmnu_VacationIncr_Err:
    glbFrmCaption$ = "Vacation Increase Rule"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "VacationIncr", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_File_Payroll_Matrix_Click()

Dim SQLQ As String
On Error GoTo GPayroll_Err

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Matrix Then
        Screen.MousePointer = HOURGLASS
        Load frmPAYROLL
        'SQLQ = "SELECT * FROM HRMATRIX"
        'frmPAYROLL.Data1.RecordSource = SQLQ
        'frmPAYROLL.Data1.Refresh
        'frmPAYROLL.Show
        frmPAYROLL.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
GPayroll_Err:
    glbFrmCaption$ = "Get Payroll"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Payroll", "SELECT")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_AttendCode_Matrix_Click()
On Error GoTo AttCodeMatrix_Err

    Screen.MousePointer = HOURGLASS
    Load frmAttendMatrix
    frmAttendMatrix.ZOrder 0
    Screen.MousePointer = DEFAULT
    
Exit Sub
AttCodeMatrix_Err:
    glbFrmCaption$ = "Attendance Code Matrix"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Attend Code Matrix", "Load")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub

Private Sub mmnu_FollowUpCodeEmail_Matrix_Click()
On Error GoTo FollowUpCodeEmail_Matrix_Err

    Screen.MousePointer = HOURGLASS
    Load frmFollowUpEMailMatrix
    frmFollowUpEMailMatrix.ZOrder 0
    Screen.MousePointer = DEFAULT
    
Exit Sub
FollowUpCodeEmail_Matrix_Err:
    glbFrmCaption$ = "Follow Up Code Email Matrix"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Follow Up Code Email Matrix", "Load")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub
 
Private Sub mmnu_DeptGL_Matrix_Click()
On Error GoTo DeptGL_Matrix_Err

    Screen.MousePointer = HOURGLASS
    Load frmDeptGLMatrix
    frmDeptGLMatrix.Caption = lStr("Department") & "/" & lStr("G/L") & " Matrix"
    frmDeptGLMatrix.ZOrder 0
    Screen.MousePointer = DEFAULT
    
Exit Sub
DeptGL_Matrix_Err:
    glbFrmCaption$ = "Department/GL Matrix"
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "frmMDI", "Department/GL Matrix", "Load")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    End If
End Sub
 
Private Sub mmnu_Find_Click()
    Call GET_EMP
End Sub

Private Sub mmnu_FindCandi_Click() 'Ticket #24184 Franks 09/11/2013
    frmSFFind.Show
End Sub

Private Sub mmnu_Followups_Click()
    MDIMain.panHelp(0).Caption = "Select a menu item"
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(3).Caption = ""
End Sub

Private Sub mmnu_Help_Click()
   ' Const HELP_KEY = &H101
   ' Me.CMDialog1.HelpFile = gflHelp$   ' Specify the Help file to open.
  '  Me.CMDialog1.HelpCommand = HELP_KEY    ' When WINHELP.EXE is executed, Help for a specified keyword will be displayed.
  '  Me.CMDialog1.HelpKey = "INFOHR_Welcome" ' Specify the keyword.
  '  Me.CMDialog1.Action = 6    ' Execute WINHELP.EXE.
  SendKeys "{F1}"
End Sub

Private Sub mmnu_Clear_Accrual_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUAccrClr
        frmUAccrClr.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_holiday_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Holiday Then
        Screen.MousePointer = HOURGLASS
        Load frmSHoliday
        frmSHoliday.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_HrsEnt_Click()
    'If gSec_Upd_Hrly_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Hrly_Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSHrsEnt
        frmSHrsEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_HrsBasedEnt_Click()
    'If gSec_Upd_Hrly_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Hrly_Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSHrsHrlyEnt
        frmSHrsHrlyEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

'Hemu - 08/13/2003 Begin - still to add Security rights
'Jaddy changed to use Entitlements Security
Private Sub mmnu_ZeroOutEnt_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Unload frmUEntitle
        Load frmUEntitle
        
        'Ticket #17924 - Begin
        'frmUEntitle.cmdRolloverHourly.Visible = False  'Menu Item added
        'frmUEntitle.cmdZeroOutHourly.Visible = True    'Menu Item added
        'Ticket #17924 - End
        
        frmUEntitle.cmdZeroOut_Click
        frmUEntitle.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_RollOverEnt_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Unload frmUEntitle
        Load frmUEntitle
        
        'Ticket #17924 - Begin
        'frmUEntitle.cmdRolloverHourly.Visible = True   'Menu Item added
        'frmUEntitle.cmdZeroOutHourly.Visible = False   'Menu Item added
        'Ticket #17924 - End
        
        frmUEntitle.cmdRollover_Click
        frmUEntitle.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
'Hemu - 08/13/2003 End

Private Sub mmnu_ZeroOutHrEnt_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUHrsEnt
        frmUHrsEnt.Caption = "Zero Out Hourly Entitlement"
        frmUHrsEnt.cmdZeroOutHr_Click
        frmUHrsEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_RollOverHrEnt_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUHrsEnt
        frmUHrsEnt.Caption = "Rollover Hourly Entitlement"
        frmUHrsEnt.cmdRolloverHr_Click
        frmUHrsEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_HTile_Click()
    MDIMain.Arrange 1
End Sub

Private Sub mmnu_IMP_Photo_Click()
    If gSec_Upd_Basic Then 'ticket #17503
        FRMINPHOTO.Show
    Else
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mmnu_EmpFlags_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Setup Employee Flags"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub


Private Sub mmnu_Label_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label of Code and Date Shown"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label1_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Demographics screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label2_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Status/Dates screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label3_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Dependents screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label4_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Banking Information screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label5_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Other Information screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label6_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Employee Flags screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label20_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - Additional Payroll ID Data screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label7_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Position screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label8_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Salary screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label9_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Performance screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label10_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Attendance screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label11_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - Associations screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
            
Private Sub mmnu_Label12_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Continuing Education screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label13_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - User Defined Table screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
            
Private Sub mmnu_Label14_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - Follow-ups screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
            
Private Sub mmnu_Label15_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - Counseling screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
            
Private Sub mmnu_Label16_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - Comments screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label21_Click() 'Ticket #26254 Franks 12/09/2014
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Job Master screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Label17_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Position Master screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label18_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = "Label - Dashboard Setup screen"
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Label19_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Label Then
        Screen.MousePointer = HOURGLASS
        Unload frmSLabel
        Load frmSLabel
        frmSLabel.Caption = lStr("Label - Province/State Master screen")
        frmSLabel.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_MarketLine_Click()
    Screen.MousePointer = HOURGLASS
    Load frmMarketLine
    frmMarketLine.Show 1
    'frmMarketLine.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_JobFamily_Click() 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
        Screen.MousePointer = HOURGLASS
        Call Get_JobFamily(True, "JOBFAMILY")
        Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_SubJobFamily_Click() 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
        Screen.MousePointer = HOURGLASS
        Call Get_JobFamily(True, "SUBFAMILY")
        Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_GroupJob_Click() 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
        Screen.MousePointer = HOURGLASS
        Call Get_JobFamily(True, "GROUPJOBS")
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Lgr_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Ledgers Then
        Screen.MousePointer = HOURGLASS
        Call Get_Ledgers(True)
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_EnterLeave_Click()
        Screen.MousePointer = HOURGLASS
        Load frmUEnterLeave
        frmUEnterLeave.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Mass_Attendance_Click()
    'If gSec_Upd_Attendance Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Attendance_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUATTEND
        frmUATTEND.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_AttHis_Click()
    'If gSec_Upd_Attendance Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Attendance_His_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUATTHIS
        frmUATTHIS.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Benefits_Click()
    'If gSec_Upd_Benefits Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Benefits_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUBENEFITS
        frmUBENEFITS.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_CDE_Click()
    'If gSec_Upd_Other_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Other_Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmDolEntit
        frmDolEntit.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_DoorAccess_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_DoorAccess Then
        Screen.MousePointer = HOURGLASS
        Load frmUDOORS
        frmUDOORS.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Code_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Mass_Codes Then
        Load frmUCode
        frmUCode.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_COE_Click()
    'If gSec_Upd_Earnings Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Other_Earnings_MassUpdate", glbUserID) Then
        frmUOtherEarn.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

'Private Sub mmnu_Mass_ConvertDatabase_Click()
'    frmImportDb.Show vbModal
'End Sub

Private Sub mmnu_mass_EducSemin_Click()
    'If gSec_Upd_Education_Seminars Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Education_Seminars_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUSEMINARS
        frmUSEMINARS.ZOrder 0
       ' Load frmEseminars
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_ImportAttachment_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("ImpAttachment_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUImpAttachFile
        frmUImpAttachFile.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_DocTypeInfoUpdate_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("ImpAttachment_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUAssignDocType
        frmUAssignDocType.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Followup_Click()
    'If gSec_Upd_Follow_Ups Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Follow_Ups_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUFollow
        frmUFollow.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_term_Num_Change_Click()
    'If gSec_Upd_Basic And gSec_Upd_Salary Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Call UnloadFrms
        Load frmTermUEmpNum
        frmTermUEmpNum.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Num_Change_Click()
    'If gSec_Upd_Basic And gSec_Upd_Salary Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Call UnloadFrms
        Load frmUEmpNum
        frmUEmpNum.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Position_Click()
    'If gSec_Upd_Position And gSec_Upd_Job_Master And gSec_Upd_Salary Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Job_Master_MassUpdate", glbUserID) Then
        frmUJobs.Show '1
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_EmployeePosition_Click()
    frmUEmpPos.Show
End Sub

Private Sub mmnu_Mass_ReportAuth_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_Position And gSec_Inq_Performance Then
        frmURepAuth.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Salary_Click()
    'If gSec_Upd_Salary Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Salary_MassUpdate", glbUserID) Then
        frmUSalary.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_TD1Dollar_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_Basic Then
        frmUTd1Dollar.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_Terminations_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_Terminations Then
        frmUTERM.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Mass_WorkSchedule_Click()
    frmUScheduler.Show
End Sub

Private Sub mmnu_Mass_ImportEmailAddress_Click()
    frmUEmailLoad.Show
End Sub

Private Sub mmnu_New_Hire_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_New_Hire Then
        Load frmNewHire
        frmNewHire.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Occ_Class_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Classes Then
        glbOClassMode% = False   ' are we in select mode
        panHelp(0).Caption = "info:HR Main functions Locked until exit this form."
        Screen.MousePointer = HOURGLASS
        Load frmOCCLASS
        Screen.MousePointer = DEFAULT
        frmOCCLASS.Show 1
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Opus_Payroll_Click()
    Load frmOpus
    frmOpus.ZOrder 0
End Sub

Private Sub mmnu_PayWeb_Exp_Attd_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Attendance Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /exportattendance /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /exportattendance", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
    

Private Sub mmnu_PayWeb_Exp_IDL_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /exportidl /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /exportidl", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayWeb_Exp_Ongoing_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /exportongoing /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /exportongoing", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Payweb_Reset_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If Not gSec_Upd_Audit Then
    '    MsgBox "You Do Not Have Authority For This Transacaction"
    '    Exit Sub
    'Else
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /resetflag /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /resetflag", vbMaximizedFocus
        End If
    'End If
End Sub

Private Sub mmnu_PayWeb_Imp_Attd_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Attendance Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /importattendance /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /importattendance", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayWeb_Imp_IDL_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Employee And gSec_Import_Salaries And gSec_Import_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /importidl /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /importidl", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayWeb_Imp_Ongoing_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Employee And gSec_Import_Salaries And gSec_Import_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /importongoing /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /importongoing", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayWeb_Imp_YTD_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Employee And gSec_Import_Salaries And gSec_Import_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /importytd /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /importytd", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayWeb_Setup_Click()
    If InStr(LCase(Command), "/path") > 0 Then
        Shell glbPayWebEXE & " /setup /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
    Else
        Shell glbPayWebEXE & " /setup", vbMaximizedFocus
    End If
End Sub

Private Sub mmnu_PayWeb_Code_Matrix_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Matrix Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbPayWebEXE & " /codematrix /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbMaximizedFocus
        Else
            Shell glbPayWebEXE & " /codematrix", vbMaximizedFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Accrual_Class_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Attendance Or gSec_Export_Attendance Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /accrualclass /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /accrualclass", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Code_Sync_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /codesync /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /codesync", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_IHRCode_Sync_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /ihrcodesync /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /ihrcodesync", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Att_Sync_Click()
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Attendance Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /att_sync /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /att_sync", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_IDL_Accural_Click()
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Import_Attendance Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /idl_acc /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /idl_acc", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Import_Table_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /importtable /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /importtable", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Code_Matrix_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /codematrix /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /codematrix", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PayMatrixBenefit_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /paymatrixbenefit /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /paymatrixbenefit", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Pay_Code_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /paycodemaster /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /paycodemaster", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Setup_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /vadimsetup /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /vadimsetup," & glbUserID, vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Report_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /vadimreport /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /vadimreport", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Salary_Report_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /vadimsalreport /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /vadimsalreport", vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Vadim_Database_Setup_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then
        If InStr(LCase(Command), "/path") > 0 Then
            Shell glbVadimEXE & " /databasesetup /path " & Mid(Command, InStr(LCase(Command), "/path") + 6), vbNormalFocus
        Else
            Shell glbVadimEXE & " /databasesetup," & glbUserID, vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Integration_Click(Product_Info As String, xFunction)
Dim CommandStr
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Attendance And gSec_Import_Attendance Then
        If Product_Info = "Advanced Tracker" Then
            CommandStr = "Advanced Tracker"
            CommandStr = CommandStr & "," & xFunction
            If Len(glbPlantCode) > 0 Then
                CommandStr = CommandStr & "/" & glbPlantCode
            Else
                CommandStr = CommandStr & "/ALL"
            End If
            Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
        Else
            Shell glbIntegrationEXE & " " & Product_Info & "," & xFunction, vbNormalFocus
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_GP_ImportAttendance_Click(Product_Info As String)
Dim CommandStr
    If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
        '2172W - County of Lanark
        'Ticket #19782 Franks 02/02/2011 for Frontenac
        MsgBox "Not part of your integration"
    Else
        If gSec_Import_Attendance Then
            'Shell glbIntegrationEXE & " " & Product_Info & ",Import Attendance", vbNormalFocus
            CommandStr = Product_Info & ",Import Attendance"
            If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #22342
                If Len(glbPlantCode) > 0 Then
                    CommandStr = CommandStr & ",," & glbPlantCode
                Else
                    'CommandStr = CommandStr & ",,ALL"
                    MsgBox "You do not have security setup for one " & lStr("Division") & " " & Chr(10) & "Cannot do Attendance Import "
                    Exit Sub
                End If
            End If
            Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
    End If
End Sub

Private Sub mmnu_GP_ImportEntAtt_Click(Product_Info As String)
    If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
        '2172W - County of Lanark
        'Ticket #19782 Franks 02/02/2011 for Frontenac
        If gSec_Import_Attendance Then
            Shell glbIntegrationEXE & " " & Product_Info & ",Import Attendance,EntAtt", vbNormalFocus
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
    Else
        MsgBox "Not part of your integration"
    End If
End Sub

Private Sub mmnu_GP_ImportAttIDL_Click(Product_Info As String)
    If glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2410W" Then
        '2172W - County of Lanark
        'Ticket #19782 Franks 02/02/2011 for Frontenac
        If gSec_Import_Attendance Then
            Shell glbIntegrationEXE & " " & Product_Info & ",Import Attendance,AttIDL", vbNormalFocus
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
    Else
        MsgBox "Not part of your integration"
    End If
End Sub

Private Sub mmnu_GP_IncomeCodeMatrix_Click(Product_Info As String)
    If glbCompSerial = "S/N - 2182W" Then 'Town of Caledon - Ticket #25024 Franks 11/22/2014
        MsgBox "Not part of your integration"
    Else
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then
            Shell glbIntegrationEXE & " " & Product_Info & ",IncomeCodeMatrix", vbNormalFocus
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
    End If
End Sub

Private Sub mmnu_Other_Data_Setup_Click(Product_Info As String)
'If glbWFC Then
'    If gSec_WFC_Bonus_Intergration_Interface Then
'        Shell glbIntegrationEXE & " " & Product_Info & ",Database Setup", vbNormalFocus
'    Else
'        MsgBox "You Do Not Have Authority For This Transaction"
'    End If
'Else

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    If Not glbCwis Then 'And gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then
        Shell glbIntegrationEXE & " " & Product_Info & ",Database Setup", vbNormalFocus
    
    ElseIf Product_Info = "CWIS" Then 'Simona- Leeds and Grenville ticket#14890
        frmOtherDatabaseSetup.Product_Info = Product_Info
        Load frmOtherDatabaseSetup
        frmOtherDatabaseSetup.Show 1
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mmnu_Time_Bank_Sync_Click(Product_Info As String)
Dim CommandStr
    CommandStr = "Advanced Tracker"
    CommandStr = CommandStr & "," & "TimeBank Sync"
    If Len(glbPlantCode) > 0 Then
        CommandStr = CommandStr & "/" & glbPlantCode
    Else
        CommandStr = CommandStr & "/ALL"
    End If
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End Sub

Private Sub mmnu_Adv_ExpVacSickBalance_Click(Product_Info As String)
    Shell glbIntegrationEXE & " " & Product_Info & ",VacSickBal", vbNormalFocus
End Sub
Private Sub mmnu_Adv_ExpEmpTblIDL_Click(Product_Info As String)
    Shell glbIntegrationEXE & " " & Product_Info & ",IDL", vbNormalFocus
End Sub

Private Sub mmnu_GP_HOLDING_Click(Product_Info As String)
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then
        Shell glbIntegrationEXE & " " & Product_Info & ",Holding File", vbNormalFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Integraion_Setup_Click(Product_Info As String)
'If glbWFC Then
'    If gSec_WFC_Bonus_Intergration_Interface Then
'        Shell glbIntegrationEXE & " " & Product_Info & ",Integration Setup", vbNormalFocus
'    Else
'        MsgBox "You Do Not Have Authority For This Transaction"
'    End If
'Else
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then
        Shell glbIntegrationEXE & " " & Product_Info & ",Integration Setup", vbNormalFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
'End If
End Sub

Private Sub mmnu_Integration_Selection_Click(Product_Info As String)
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then
        Shell glbIntegrationEXE & " " & Product_Info & ",Integration Selection", vbNormalFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Pos_BAND_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_SalaryGrids Then
        Load frmMBand
        frmMBand.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_IncentivePlan_CreateSpreadsheet_Click()
        Screen.MousePointer = HOURGLASS
        glbWFC_IPPopFormName = "WFCIPSpreadSheetCreate"
        Unload frmIPCreateSheet
        Load frmIPCreateSheet
        frmIPCreateSheet.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_UptOtherEarnings_Click()
        Screen.MousePointer = HOURGLASS
        glbWFC_IPPopFormName = "WFCIPUptOtherEarnings"
        Unload frmIPCreateSheet
        Load frmIPCreateSheet
        frmIPCreateSheet.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_PrintSpreadsheet_Click()
        'Ticket #29810 Franks 02/27/2017 -
        '2/13/2017 10:16:47 AM - Incoming Call
        'Jerry called, he had a call with MZ in last Friday, Peter sent the IP Excel files to the plants, and they made lots of changes on the original files, so they can't import the Excel file back to the info:HR database
        'then the program can not print it since there is no IP data in the database
        Exit Sub
        
        Screen.MousePointer = HOURGLASS
        glbWFC_IPPopFormName = "WFCIPPrintSpreadsheet"
        Unload frmIPCreateSheet
        Load frmIPCreateSheet
        frmIPCreateSheet.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_PrintEmpLetter_Click()
        Screen.MousePointer = HOURGLASS
        glbWFC_IPPopFormName = "WFCIPPrintEmpLetter"
        Unload frmIPCreateSheet
        Load frmIPCreateSheet
        frmIPCreateSheet.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_PreparePayroll_Click()
        Screen.MousePointer = HOURGLASS
        glbWFC_IPPopFormName = "WFCIPPreparePayroll"
        Unload frmIPCreateSheet
        Load frmIPCreateSheet
        frmIPCreateSheet.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_ImpSpreadsheet_Click()
        Screen.MousePointer = HOURGLASS
        glbWFC_IPPopFormName = "WFCIPSpreadSheetImport"
        Unload frmIPCreateSheet
        Load frmIPCreateSheet
        frmIPCreateSheet.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_Factors_Click()
        Screen.MousePointer = HOURGLASS
        Load frmIPFactors
        frmIPFactors.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_IncentivePlan_ImpCurrency_Click()
        glbWFC_IPPopFormName = "ImpCurrency"
        Unload frmCheckListView
        frmCheckListView.Show 1
        UnloadFrms
End Sub
Private Sub mmnu_IncentivePlan_CurrencyTable_Click()
        Screen.MousePointer = HOURGLASS
        Load frmIPCurrencyExch
        frmIPCurrencyExch.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Pos_Skills_Click()
Dim fglbEditMode%, xPos

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master And gSec_Inq_Job_Skills Then
        'If Not frmMPOSITIONS.Data1.Recordset.EOF And Not frmMPOSITIONS.Data1.Recordset.BOF Then
            If fglbEditMode Then
                MsgBox "Changes pending - save or cancel first"
                Exit Sub
            End If
            xPos = Not Len(glbPos$) = 0
    '        glbPos$ = frmMPOSITIONS.txtPosition
    '        glbPosDesc$ = frmMPOSITIONS.txtPosDescr
            'Unload frmMPOSITIONS
            Screen.MousePointer = HOURGLASS
            Load frmPosSkills
            frmPosSkills.ZOrder 0
            If xPos Then frmPosSkills.ZOrder 0
            Screen.MousePointer = DEFAULT
        'Else
        '    MsgBox "No positions to select from"
        'End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_AppTrack_LetterByPosType_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSLetterPosType
    frmSLetterPosType.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_AppTrack_AppFormWorkflow_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSAppFrmWorkflow
    frmSAppFrmWorkflow.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_AppTrack_AppFormDefaults_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSAppFormDefaults
    frmSAppFormDefaults.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Positions_Click()
    MDIMain.panHelp(0).Caption = "Select a menu item"
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(3).Caption = ""
End Sub

Private Sub mmnu_Prov_Master_Click()
    Call Get_Prov(True) ' master call of Provinces
End Sub

Private Sub mmnu_Help_Desc_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_HelpDescSetup Then
        Screen.MousePointer = HOURGLASS
        frmMHelp.Show 1
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_PosGrp_PerfCat_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Performance And gSec_Inq_Job_Master Then
        Screen.MousePointer = HOURGLASS
        frmPGrpPerfCatLnk.Show 1
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Employee_Type_Matrix_Click()
    Screen.MousePointer = HOURGLASS
    frmSEmpTypeMatrix.Show 1
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_Person_Completing_Form7_Click()
    Screen.MousePointer = HOURGLASS
    frmSPersCompltgF7.Show 1
    Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_Root_Cause_Event_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        frmEHSCauseLinks.LinkItem = "EVENT"
        frmEHSCauseLinks.Show 1
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Table_Master_CodeLinks_Click() 'Ticket #21106 Franks 11/04/2011
        Screen.MousePointer = HOURGLASS
        frmTABLLINKS.Show 1
        Screen.MousePointer = DEFAULT
End Sub
Private Sub mmnu_Root_Cause_Immediate_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        frmEHSCauseLinks.LinkItem = "IMMEDIATE"
        frmEHSCauseLinks.Show 1
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Root_Basic_Underlying_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        frmEHSCauseLinks.LinkItem = "BASIC"
        frmEHSCauseLinks.Show 1
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_AccCost_Click()
    Load frmRINCOST
    frmRINCOST.ZOrder 0
End Sub

Private Sub mmnu_R_Associations_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Associations Then
        FormAssoc% = True
        If FormDoll% = True Or FormEduc% = True Or FormOther% = True Then
            FormDoll% = False
            FormEduc% = False
            FormOther% = False
            Unload frmRAssoc
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = lStr("Associations") & " Report"  'Laura 20 Oct, 1997
        Else
            Load frmRAssoc
            frmRAssoc.Caption = lStr("Associations") & " Report"
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_AttdCost_Click()
'Changed by Bryan 26/Apr/06 Ticket#10730
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Cost_Of_Employment And gSec_Inq_Attendance_History Then 'gSec_Rpt_Master_Attendance
        If frmRAttend.Visible Then
            Unload frmRAttend
            frmRAttend.Caption = "Costed Attendance Report"
            Call frmRAttend.comGrpLoad
            frmRAttend.ZOrder 0
        Else
            frmRAttend.Caption = "Costed Attendance Report"
            Load frmRAttend
            frmRAttend.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_AttdHist_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Attendance_History Then  'gSec_Rpt_Master_Attendance Then
        If frmRAttend.Visible Then
            Unload frmRAttend
            frmRAttend.Caption = "Attendance Report Including Historical Data"
            Call frmRAttend.comGrpLoad
            frmRAttend.ZOrder 0
        Else
            frmRAttend.Caption = "Attendance Report Including Historical Data"
            Load frmRAttend
            frmRAttend.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EnviroServ_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Attendance_History Then  'gSec_Rpt_Master_Attendance Then
        If frmRAttend.Visible Then
            Unload frmRAttend
            frmRAttend.Caption = "Wellington Terrace Attendance Report"
            Call frmRAttend.comGrpLoad
            frmRAttend.ZOrder 0
        Else
            frmRAttend.Caption = "Wellington Terrace Attendance Report"
            Load frmRAttend
            frmRAttend.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_ESSReqTrnAudit_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Attendance_History Then  'gSec_Rpt_Master_Attendance Then
        If frmRESSTrnAudit.Visible Then
            Unload frmRESSTrnAudit
            frmRESSTrnAudit.Caption = "ESS Requests - Transaction Audit Report"
            frmRESSTrnAudit.ZOrder 0
        Else
            frmRESSTrnAudit.Caption = "ESS Requests - Transaction Audit Report"
            Load frmRESSTrnAudit
            frmRESSTrnAudit.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_AttWrkSchDescrepancy_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        If frmRAttDiscrepancy.Visible Then
            frmRAttDiscrepancy.Caption = "Attendance/Work Schedule Discrepancy"
            frmRAttDiscrepancy.ZOrder 0
        Else
            frmRAttDiscrepancy.Caption = "Attendance/Work Schedule Discrepancy"
            Load frmRAttDiscrepancy
            frmRAttDiscrepancy.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_AttdPersonalDayRpt_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        If frmRAttend.Visible Then
            Unload frmRAttend
            frmRAttend.Caption = "Personal Day Report"
            Call frmRAttend.comGrpLoad
            frmRAttend.ZOrder 0
        Else
            frmRAttend.Caption = "Personal Day Report"
            Load frmRAttend
            frmRAttend.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Attdpoint_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        frmRPoint.Caption = "Attendance Bonus Points Report"
        Load frmRPoint
        frmRPoint.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Attendance_Calendar_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        frmRAttSht.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Accrual_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Entitlements Then
        Load frmRAccrual
        frmRAccrual.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Attendance_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        If frmRAttend.Visible Then
            Unload frmRAttend
            frmRAttend.Caption = "Attendance Report"
            Call frmRAttend.comGrpLoad
            frmRAttend.ZOrder 0
        Else
            frmRAttend.Caption = "Attendance Report"
            Load frmRAttend
            frmRAttend.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Benefit_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Benefits Then
        Load frmRBenefits
        frmRBenefits.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PosTablesExp_Click()
    Unload frmRPosBudWFC
    glbFrmCaption$ = "Position Control Table Exports"
    Load frmRPosBudWFC
    frmRPosBudWFC.ZOrder 0
End Sub
Private Sub mmnu_R_BudPos_Click()
    Unload frmRPosBudWFC
    glbFrmCaption$ = "Budgeted Position Report"
    Load frmRPosBudWFC
    frmRPosBudWFC.ZOrder 0
End Sub

Private Sub mmnu_R_Birthday_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Age Then
        If frmRDOB.Visible Then
            Unload frmRDOB
            Load frmRDOB
            frmRDOB.Caption = "Birthday/Age Report"
        Else
            Load frmRDOB
            frmRDOB.ZOrder 0
            frmRDOB.Caption = "Birthday/Age Report"
        End If
        frmRDOB.lblGrp(1).Visible = True
        frmRDOB.comGroup(1).Visible = True
    '    frmRDOB.lblMonth.Caption = "Month of Birth"
        frmRDOB.lblYear.Visible = True
        frmRDOB.txtYear.Visible = True
        frmRDOB.lblMonth.Visible = True
        frmRDOB.txtMonth.Visible = True
    '    frmRDOB.txtEEID.Visible = False
    '    frmRDOB.lblEEName.Visible = False
    '    frmRDOB.comGroup(2).Top = 3750
    '    frmRDOB.lblGrp(2).Top = 3750
            
    'Hemu
        frmRDOB.chkSupBirthYrAge.Visible = gSec_Show_DOB    'Ticket #23772
    'Hemu
            
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        frmRDOB.clpDiv = ""
        frmRDOB.clpDept = ""
        frmRDOB.clpCode(1) = ""
        frmRDOB.clpCode(2) = ""
        frmRDOB.clpPT = ""
        frmRDOB.txtMonth = ""
        frmRDOB.txtYear = ""
        frmRDOB.elpEEID = ""
        frmRDOB.clpDiv.SetFocus
    
        frmRDOB.comGroup(0).Clear
        frmRDOB.comGroup(1).Clear
        frmRDOB.comGroup(2).Clear
        If frmRDOB.Caption = "Birthday/Age Report" Then
            frmRDOB.comGroup(1).AddItem "Employee Name"
            frmRDOB.comGroup(1).AddItem "Year of Birth"
            frmRDOB.comGroup(1).AddItem "Month of Birth"
            frmRDOB.comGroup(1).AddItem "(none)"
            
            frmRDOB.comGroup(2).AddItem "Employee Number"
            frmRDOB.comGroup(2).AddItem "Month of Birth"
            
            frmRDOB.comGroup(1).ListIndex = 0
            frmRDOB.comGroup(2).Enabled = True
        Else
            frmRDOB.comGroup(2).AddItem "Employee Name"
            frmRDOB.comGroup(2).AddItem "Employee Number"
            frmRDOB.comGroup(2).Enabled = True
        End If
        frmRDOB.comGroup(2).ListIndex = 0
        frmRDOB.comGroup(0).AddItem lStr("Division")
        frmRDOB.comGroup(0).AddItem lStr("Department")
        frmRDOB.comGroup(0).AddItem lStr("Location")
        frmRDOB.comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
        frmRDOB.comGroup(0).AddItem lStr("Region")
        
        If glbLinamar Then ' Frank May 2,2001
            frmRDOB.comGroup(0).AddItem "Employment Type"
            frmRDOB.comGroup(0).AddItem ("Home Line")
        End If
        If frmRDOB.Caption = "Birthday/Age Report" Then
            frmRDOB.comGroup(0).AddItem "Year of Birth"
            'Hemu - 07/10/2003 Begin
            frmRDOB.comGroup(0).AddItem "Month of Birth"
            'Hemu - 07/10/2003 End
        End If
        
        If Not glbMulti Then frmRDOB.comGroup(0).AddItem "Shift"
      
        frmRDOB.comGroup(0).AddItem "(none)"
        frmRDOB.comGroup(0).ListIndex = 0
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_C_Total_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False

    'laura dec 04, 1997
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Heatlh_Safety Then
        If Not frmRWCBCS.Visible Then
          Load frmRWCBCS
        End If
          frmRWCBCS.Caption = "Total Cost Report"
          frmRWCBCS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

    'If gSec_Rpt_Heatlh_Safety Then
    '    Load frmRINCS
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_C_WCBI_Click()
    'employee cost summary
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Heatlh_Safety Then
        If Not frmRWCBCS.Visible Then
          Load frmRWCBCS
        End If
        frmRWCBCS.Caption = "Employee/WSIB Cost Report"
        frmRWCBCS.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
    
    'If gSec_Rpt_Heatlh_Safety Then
    '    Load frmRWCBCS
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_C_CompanyAssoc_Cost_Click()
    'Company Associated Cost Report
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Heatlh_Safety Then
        If Not frmRCompCost.Visible Then
          Load frmRCompCost
        End If
        frmRCompCost.Caption = "Company Associated Cost Report"
        frmRCompCost.ZOrder 0
End Sub

Private Sub mmnu_R_OvertimeBank_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Ovt_Bank Then
        If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
            frmROvtBank.Caption = "Extra Time Bank Report"
        Else
            frmROvtBank.Caption = "Overtime Bank Report"
        End If
        Load frmROvtBank
        frmROvtBank.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_OvertimeLostHours_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Ovt_Lost_Hours Then
        If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
            frmROvtLostHrs.Caption = "Extra Time Bank Lost Hours Report"
        Else
            frmROvtLostHrs.Caption = "Overtime Bank Lost Hours Report"
        End If
        Load frmROvtLostHrs
        frmROvtLostHrs.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Compensatory_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Compensatory_Time Then
        If frmRCompTime.Visible Then
            If glbCompSerial = "S/N - 2425W" Then   'Four Villages - Ticket #19998
                frmRCompTime.Caption = "Extra Time Report"
            Else
                frmRCompTime.Caption = "Compensatory Time Report"
            End If
            Call frmRCompTime.comGrpLoad
            frmRCompTime.ZOrder 0
        Else
            If glbCompSerial = "S/N - 2425W" Then   'Four Villages - Ticket #19998
                frmRCompTime.Caption = "Extra Time Report"
            Else
                frmRCompTime.Caption = "Compensatory Time Report"
            End If
            Load frmRCompTime
            frmRCompTime.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_FlexBank_Click()
    If frmRCompTime.Visible Then
        frmRCompTime.Caption = "Flex Bank Report"
        Call frmRCompTime.comGrpLoad
        frmRCompTime.ZOrder 0
    Else
        frmRCompTime.Caption = "Flex Bank Report"
        Load frmRCompTime
        frmRCompTime.ZOrder 0
    End If
End Sub

Private Sub mmnu_R_ComPlan_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Salaries Then
        Load frmRComPlan
        frmRComPlan.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_CostER_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Cost_Of_Employment Then
        Load frmRCostE
        frmRCostE.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Counsel_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Counselling Then
        Load frmRCounsel
        frmRCounsel.ZOrder 0
    'Else
    '    MsgBox "You do not have authority for this transaction."
    'End If
    Exit Sub
End Sub

Private Sub mmnu_R_DocumentType_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_DocumentType Then
        Load frmRDocType
        frmRDocType.ZOrder 0
    'Else
    '    MsgBox "You do not have authority for this transaction."
    'End If
    Exit Sub
End Sub

Private Sub mmnu_R_CustomReport_Click()
    frmRCsRpt.Show
    frmRCsRpt.ZOrder 0
End Sub

Private Sub mmnu_R_HCASCustomReport_Click()
    frmRHCASCsRpt.Show
    frmRHCASCsRpt.ZOrder 0
End Sub

Private Sub mmnu_R_Dependents_Click()
    'laura Oct 28, 1997
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Dependents Then
        FormDepend% = True
        If FormHomeAddress% = True Then
            FormHomeAddress% = False
            Unload frmRHomeA
            Load frmRHomeA
            frmRHomeA.ZOrder 0
            frmRHomeA.Caption = "Dependents Report"
        Else
            Load frmRHomeA
            frmRHomeA.ZOrder 0
            frmRHomeA.Caption = "Dependents Report"
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_DolEnt_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_DolEnt Then
            '~~~~~~~~~~~~~     Laura
        FormDoll% = True
        If FormAssoc% = True Or FormEduc% = True Or FormOther% = True Then 'laura
            FormAssoc% = False  'Laura
            FormEduc% = False   'Laura
            FormOther% = False   'Laura
            
            Unload frmRAssoc
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = "Dollar Entitlement Report"  'Laura 20 Oct, 1997
        Else
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = "Dollar Entitlement Report"
        End If
        '~~~~~~~Laura
       '' Load frmRAssoc
    '' Load frmRDolEnt
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_DoorAccess_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_DoorAccess Then
        Load frmRDoors
        frmRDoors.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Education_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Education_Seminars Then
        Load frmREdSem
        frmREdSem.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EELabels_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Employee_Labels Then
        Load frmRLabels
        frmRLabels.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEFlags_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Employee_Flags Then
        Load frmREmpFlags
        frmREmpFlags.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEMaster_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Profiles Or gSec_Inq_Comments Then
        Load frmRMaster
        frmRMaster.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEGLDistribution_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_GLDistribution Then
        Load frmRGLDistribution
        frmRGLDistribution.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEmergLeave_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Emergency_Leave Then
        Load frmREmergLeave
        frmREmergLeave.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEPosition_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Job_List Then
        Unload frmRPosition
        FormEmplPosition% = True
        FormLanguages% = False
        If frmRPosition.Visible = True Then
            Unload frmRPosition
        End If
        frmRPosition.Caption = "Employee/Position Report"
        Load frmRPosition
        frmRPosition.ZOrder 0
        'frmRPosition.panReportG.Visible = True 'js-31Mar99
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEHistory_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_EMP_HISTORY Then
        Unload frmRHistory
        'FormEmplHistory% = True
        If frmRHistory.Visible = True Then
            Unload frmRHistory
        End If
        frmRHistory.Caption = "Employee History Report"
        Load frmRHistory
        frmRHistory.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EESN2343_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Job_List Then
        Unload frmRPosition
        FormEmplPosition% = True
        FormLanguages% = False
        If frmRPosition.Visible = True Then
          Unload frmRPosition
        End If
        frmRPosition.Caption = "Category/Status Report"
        Load frmRPosition
        frmRPosition.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EEProfile_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Profiles Then
        Unload frmRPosition
        FormEmplPosition% = False   'Serbo
        FormLanguages% = False
        frmRPosition.Caption = "Employee Profile Report"
        Load frmRPosition
        frmRPosition.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Email_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Profiles Then
        Load frmREmail
        frmREmail.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Emergency_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Emergecy_Contacts Then
        Screen.MousePointer = HOURGLASS
        Load frmREmergC
        frmREmergC.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_R_Entitlements_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Entitlements Then
        Load frmREntitle
        frmREntitle.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Followup_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRFollow
        frmRFollow.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_FollowupEmailLog_Click()
    Load frmRFollowEmailLog
    frmRFollowEmailLog.ZOrder 0
End Sub

Private Sub mmnu_VacEntDailySkippedLog_Click()
    Load frmRDEntSkipLog
    frmRDEntSkipLog.ZOrder 0
End Sub

Private Sub mmnu_VacDailyAccDetails_Click()
    Load frmRDailyAccDtls
    frmRDailyAccDtls.ZOrder 0
End Sub

Private Sub mmnu_R_IWantToKnowYou_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRIWantYouToKnow
        frmRIWantYouToKnow.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_ITHire_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRITHire
        frmRITHire.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_ITNoticeOfChange_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRITNoticeOfChange
        frmRITNoticeOfChange.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_NoticeOfChange_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRNoticeOfChange
        frmRNoticeOfChange.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PerfImproveActionPlan_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRPerfImprovActPlan
        frmRPerfImprovActPlan.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PerformanceReviewForm_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRPerfReviewRpt
        frmRPerfReviewRpt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Separation_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRSeparationRpt
        frmRSeparationRpt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_TerminationForm_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRTerminationHRRpt
        frmRTerminationHRRpt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_UpdateMeeting_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRUpdMeetingRpt
        frmRUpdMeetingRpt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Warning_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Follow_Ups Then
        Load frmRWarningRpt
        frmRWarningRpt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Formal_Education_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Formal_Education Then
       ' Load frmRFormEd
                '~~~~~~~~~~~~~     Laura
            FormEduc% = True
      If FormAssoc% = True Or FormDoll% = True Or FormOther% = True Then 'laura
            FormAssoc% = False  'Laura
            FormDoll% = False   'Laura
            FormOther% = False   'Laura
            
            Unload frmRAssoc
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = "Formal Education Report"  'Laura 20 Oct, 1997
        Else
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = "Formal Education Report"
        End If
        '~~~~~~~Laura
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Home_Click()
    'If gSec_Rpt_Home_Address Then
    '    Load frmRHomeA
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
    ''frmRHomeA
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Home_Address And gSec_Show_ADDRESS Then
        FormHomeAddress% = True
        If FormDepend% = True Then
            FormDepend% = False
            Unload frmRHomeA
            Load frmRHomeA
            frmRHomeA.ZOrder 0
            frmRHomeA.Caption = "Home Address/Telephone Report"
        Else
            Load frmRHomeA
            frmRHomeA.ZOrder 0
            frmRHomeA.Caption = "Home Address/Telephone Report"
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_EmployeeDates_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Unload frmREmpDates
        Load frmREmpDates
        frmREmpDates.ZOrder 0
        frmREmpDates.Caption = "Employee Dates Report"
    
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_LengthOfService_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Unload frmRLenServ
        Load frmRLenServ
        frmRLenServ.ZOrder 0
        frmRLenServ.Caption = "Length of Service Report"
    
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_LOA_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Unload frmRLOA
        Load frmRLOA
        frmRLOA.ZOrder 0
        frmRLOA.Caption = "Leave of Absence Report"
    
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_HrEnt_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_HourEnt Then
        Load frmRHrEnt
        frmRHrEnt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Timesheet_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        Load frmRTimesheet
        frmRTimesheet.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_TimesheetWCost_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        Load frmRTimesheetWCost
        frmRTimesheetWCost.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_TimesheetStatus_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        Load frmRTSStatus
        frmRTSStatus.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_JournalEntry_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Attendance Then
        Load frmRJournalEntry
        frmRJournalEntry.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Body_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
     'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Body Site Report"
        frmRINJURY.ReportType.Caption = "1"
        frmRINJURY.Report1.Visible = True
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = False
        frmRINJURY.lblHSShift.Visible = True
        frmRINJURY.txtShift(3).Visible = True
        
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Code_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Injury Code Report"
        frmRINJURY.ReportType.Caption = "7"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = True
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = False
        frmRINJURY.lblHSShift.Visible = True
        frmRINJURY.txtShift(3).Visible = True
        
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Day_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Day of Week Report"
        frmRINJURY.ReportType.Caption = "2"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = False
        frmRINJURY.lblHSShift.Visible = True
        frmRINJURY.txtShift(3).Visible = True
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_EE_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Employee Report"
        frmRINJURY.ReportType.Caption = "3"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = False
        frmRINJURY.lblHSShift.Visible = True
        frmRINJURY.txtShift(3).Visible = True
        frmRINJURY.chkComments.Visible = True
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Experience_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Experience Report"
        frmRINJURY.ReportType.Caption = "5"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = True
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = False
        frmRINJURY.lblHSShift.Visible = True
        frmRINJURY.txtShift(3).Visible = True
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Incident_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        'If App.Path = "\\HRSS_TORONTO\PROGRAMMING\HR Systems VB6\Ihr 5.0" Or App.Path = "U:\HR SYSTEMS VB6\IHR 5.0" Then
        '    Load frmRINTYPT
        'Else
            Load frmRINTYPE
            frmRINTYPE.ZOrder 0
        'End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Plant_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Plant Area Report"
        frmRINJURY.ReportType.Caption = "8"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = True
        frmRINJURY.Report9.Visible = False
            frmRINJURY.lblHSShift.Visible = True
            frmRINJURY.txtShift(3).Visible = True
            
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Shift_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Shift Code Report"
        frmRINJURY.ReportType.Caption = "9"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = True
        frmRINJURY.lblHSShift.Visible = False
        frmRINJURY.txtShift(3).Visible = False
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_I_Trends_Click()
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Health_Safety Then
        If Not frmRINJURY.Visible Then
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        Else
            Unload frmRINJURY
            Load frmRINJURY
            frmRINJURY.ZOrder 0
        End If
        frmRINJURY.Caption = "Analysis of Incidents by Employee Trend Report"
        frmRINJURY.ReportType.Caption = "4"
        frmRINJURY.Report1.Visible = False
        frmRINJURY.Report5.Visible = False
        frmRINJURY.Report7.Visible = False
        frmRINJURY.Report8.Visible = False
        frmRINJURY.Report9.Visible = False
        frmRINJURY.lblHSShift.Visible = True
        frmRINJURY.txtShift(3).Visible = True
        
        frmRINJURY.clpDiv.Text = ""
        frmRINJURY.clpDept.Text = ""
        frmRINJURY.clpCode(1).Text = ""
        frmRINJURY.clpCode(2).Text = ""
        frmRINJURY.clpCode(3).Text = ""
        frmRINJURY.clpCode(4).Text = ""
        frmRINJURY.clpPT.Text = ""
        frmRINJURY.elpEEID.Text = ""
        frmRINJURY.dlpDateRange(0).Text = ""
        frmRINJURY.dlpDateRange(1).Text = ""
        frmRINJURY.medDOW(0).Text = ""
        frmRINJURY.medDOW(1).Text = ""
        frmRINJURY.clpDiv.SetFocus
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Languages_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Languages Then
        Unload frmRPosition
        FormLanguages% = True
        FormEmplPosition% = False
        If frmRPosition.Visible = True Then
            Unload frmRPosition
            frmRPosition.ZOrder 0
        End If
        frmRPosition.Caption = "Languages Report"
        Load frmRPosition
        frmRPosition.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_ProfitSharing_Click()
        glbFormCaption = "Profit Sharing Report"
        Unload frmRProfitSharing
        Load frmRProfitSharing
        frmRProfitSharing.ZOrder 0
End Sub
Private Sub mmnu_R_RedCircled_Click()
        glbFormCaption = "Red Circled Report"
        Unload frmRProfitSharing
        Load frmRProfitSharing
        frmRProfitSharing.ZOrder 0
End Sub

Private Sub mmnu_R_PayTrans_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayrollTrans Then
        Load frmRPayTrans
        frmRPayTrans.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_OtherEarn_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_OtherEarn Then
      '  Load frmROEarn
                '~~~~~~~~~~~~~     Laura
          FormOther% = True
        If FormAssoc% = True Or FormDoll% = True Or FormEduc% = True Then 'laura
            FormAssoc% = False  'Laura
            FormDoll% = False   'Laura
            FormEduc% = False   'Laura
            
            Unload frmRAssoc
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = "Other Earnings Report"  'Laura 20 Oct, 1997
        Else
            Load frmRAssoc
            frmRAssoc.ZOrder 0
            frmRAssoc.Caption = "Other Earnings Report"
        End If
        '~~~~~~~Laura
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Password_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Passwords Then
        Load frmRSecure
        frmRSecure.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PlanEstablishment_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then
        Load frmRPOE
        frmRPOE.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PoPage_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Key_Workforce Then
        Load frmRPoPage
        frmRPoPage.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_StaffRatios_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Staff_Management Then
        Load frmRStaffRatio
        frmRStaffRatio.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_WCLostTimeIncRate_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_WC_Time Then
        Load frmRWCIncRate
        frmRWCIncRate.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_WCLostWrkHrRate_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_WC_Work Then
        Load frmRWCHrRate
        frmRWCHrRate.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_ExternalHire_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_External_Hire Then
        If frmRExtHire.Visible Then
            Unload frmRExtHire
            Load frmRExtHire
            frmRExtHire.ZOrder 0
            frmRExtHire.Caption = "External Hire Rate Report"
            frmRExtHire.lblIntPos.Visible = False
            frmRExtHire.lblIntTrans.Visible = False
            frmRExtHire.clpCode(8).Visible = False
        Else
            frmRExtHire.Caption = "External Hire Rate Report"
            frmRExtHire.lblIntPos.Visible = False
            frmRExtHire.lblIntTrans.Visible = False
            frmRExtHire.clpCode(8).Visible = False
            Load frmRExtHire
            frmRExtHire.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_InternalHire_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Internal_Hire Then
        If frmRExtHire.Visible Then
            Unload frmRExtHire
            Load frmRExtHire
            frmRExtHire.ZOrder 0
            frmRExtHire.Caption = "Internal Transfers to Total Hires Ratio Report"
            frmRExtHire.lblIntPos.Visible = True
            frmRExtHire.lblIntTrans.Visible = True
            frmRExtHire.clpCode(8).Visible = True
        Else
            frmRExtHire.Caption = "Internal Transfers to Total Hires Ratio Report"
            frmRExtHire.lblIntPos.Visible = True
            frmRExtHire.lblIntTrans.Visible = True
            frmRExtHire.clpCode(8).Visible = True
            Load frmRExtHire
            frmRExtHire.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_TurnoverRates_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Job Then
        Load frmRATurnovrRt
        frmRATurnovrRt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PaidSickHr_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Paid_Sick Then
        Load frmRPaidSicHr
        frmRPaidSicHr.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Position_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Job Then
        Load frmRPosMaster
        frmRPosMaster.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Salary_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Salaries Then
        Load frmRSALMAST
        frmRSALMAST.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Salary_Performance_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Salary_Performance Then
        Load frmRSALPERF
        frmRSALPERF.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_PerformanceReview_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Salary_Performance Then
        Load frmRPERFORMANCEREVIEW
        frmRPERFORMANCEREVIEW.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Temp_CrossTraining_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Temp_CrossTraining Then
        Load frmRTmpCrossTrain
        frmRTmpCrossTrain.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Req_Course_Hist_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Req_Course_Hist Then
        Load frmRCourses
        frmRCourses.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_AttendaceSignIn_Click()
    Load frmRFAttSignIn
    frmRFAttSignIn.ZOrder 0
End Sub

Private Sub mmnu_R_ATTDiscipline_Click()
    Load frmRFATTDiscipline
    frmRFATTDiscipline.ZOrder 0
End Sub

Private Sub mmnu_R_COCDiscipline_Click()
    Load frmRFCOCDiscipline
    frmRFCOCDiscipline.ZOrder 0
End Sub

Private Sub mmnu_R_Seniority_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Seniority Then
        Load frmRSeniority
        frmRSeniority.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_SIN_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Basic And gSec_Show_SIN_SSN Then
        Screen.MousePointer = HOURGLASS
    '    frmRSIN.Caption = mmnu_R_SIN.Caption & " Report"
        Load frmRSIN
        frmRSIN.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Skills_Click()
    'laura oct 23, 1997
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Skills Then
        Load frmRSkills
        frmRSkills.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Succession_Click()
    'George Apr 6,2006
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Skills Then 'here need to be changed
        Load frmRSuccession
        frmRSuccession.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_GapAnalysis_Click()
    'George Apr 6,2006
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Skills Then 'here need to be changed
        Load frmRGapAnalysis
        frmRGapAnalysis.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Table_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Table_Codes Then
        Load frmTblName
        frmTblName.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Tele_Ext_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Telephone_Extensions Then
        If frmRDOB.Visible Then
            Unload frmRDOB
            Load frmRDOB
            frmRDOB.ZOrder 0
            frmRDOB.Caption = "Telephone Extension Report"
        Else
            frmRDOB.Caption = "Telephone Extension Report"
            Load frmRDOB
            frmRDOB.ZOrder 0
        End If
        frmRDOB.lblGrp(1).Visible = False
        frmRDOB.comGroup(1).Visible = False
    '    frmRDOB.lblMonth.Caption = "Employee Number" '
        frmRDOB.lblYear.Visible = False
        frmRDOB.txtYear.Visible = False
        frmRDOB.lblMonth.Visible = False
        frmRDOB.txtMonth.Visible = False
    '    frmRDOB.txtEEID.Visible = True
        frmRDOB.comGroup(2).Top = frmRDOB.comGroup(1).Top
        frmRDOB.lblGrp(2).Top = frmRDOB.lblGrp(1).Top
        
        frmRDOB.comGroup(2).AddItem "Employee Name"
        frmRDOB.comGroup(2).AddItem "Employee Number"
        frmRDOB.comGroup(2).Enabled = True
        frmRDOB.comGroup(2).ListIndex = 0
        
        frmRDOB.comGroup(0).AddItem lStr("Division")
        frmRDOB.comGroup(0).AddItem lStr("Department")
        frmRDOB.comGroup(0).AddItem lStr("Location")      'Jaddy jun 16,1998
        frmRDOB.comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
        If glbLinamar Then ' Frank May 2,2001
            frmRDOB.comGroup(0).AddItem "Employment Type"
            frmRDOB.comGroup(0).AddItem lStr("Region")
            frmRDOB.comGroup(0).AddItem ("Home Line")
        End If
        If Not glbMulti Then frmRDOB.comGroup(0).AddItem "Shift"
        frmRDOB.comGroup(0).AddItem "(none)"
        frmRDOB.comGroup(0).ListIndex = 0
      
        'Hemu
        frmRDOB.chkSupBirthYrAge.Visible = False
        'Hemu
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_WorkSchedule_Click()
    Load frmRScheduler
    frmRScheduler.ZOrder 0
End Sub

Private Sub mmnu_R_Terminations_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Termination Then
        Load frmRTerm
        frmRTerm.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_SalVacIncr_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Entitlements Then
        Load frmRSalVacPrcIncr
        frmRSalVacPrcIncr.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub
Private Sub mmnu_R_Train_Plan_Click()

        glbFormCaption = "Training Plan Report"
        Unload frmRTrainMatrix
        Load frmRTrainMatrix
        frmRTrainMatrix.ZOrder 0
End Sub
Private Sub mmnu_R_Train_Matrix_Click()
Dim fglbEditMode%, xPos
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Master_Education_Seminars Then  'Hemu - Ticket #9454 Linda asked me to remove this -> gSec_Rpt_Master_Job
        glbFormCaption = "Training Matrix Report"
        Unload frmRTrainMatrix
        Load frmRTrainMatrix
        frmRTrainMatrix.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_R_Turnover_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_Turnover Then
        Load frmRTurnov
        frmRTurnov.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Reports_Click()
    MDIMain.panHelp(0).Caption = "Select a menu item"
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(3).Caption = ""
End Sub

Private Sub mmnu_SetPerformance_Click()
On Error GoTo Err_SetPerform
    If glbSetPer = False Then Unload frmEPERFORM
    glbSetPer = True

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Performance Then
        Load frmEPERFORM
        frmEPERFORM.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_SetPerform:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_SetPosition_Click()
On Error GoTo Err_Position

    If glbSetPos = False Then Unload frmEPOSITION
    glbSetPos = True

Screen.MousePointer = HOURGLASS
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Position Then
        Load frmEPOSITION
        frmEPOSITION.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_Position:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_SetSalary_Click()
On Error GoTo Err_Salary
    
    If glbCompSerial = "S/N - 2288W" Then 'tkt#10845
        If glbSetSal = False Then Unload frmESALARYMusashi
    Else
        If glbSetSal = False Then Unload frmESALARY
    End If
    glbSetSal = True

Screen.MousePointer = HOURGLASS

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Salary Then
        If glbCompSerial = "S/N - 2288W" Then 'tkt#10845
            Load frmESALARYMusashi
            frmESALARYMusashi.ZOrder 0
        Else
            Load frmESALARY
            frmESALARY.ZOrder 0
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Screen.MousePointer = DEFAULT

Exit Sub
Err_Salary:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_Table_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table.count <> 0 Then
        Screen.MousePointer = HOURGLASS
        Call Get_Master_Code
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Table_Attendance_Group_Master_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table.count <> 0 Then
        Screen.MousePointer = HOURGLASS
        Call Get_Attendance_Group_Master_Code
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Term_EE_Click()
On Error GoTo Err_Term

    If glbWFC And glbUserID = "999999999" Then
        MsgBox "The 999999999 account can not be used for terminations.  Please log in under your employee number, and try again.", vbInformation + vbOKOnly, "Termination Not Allowed"
        Exit Sub
    End If

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        glbDivTranInPlant = "N"
        
        Unload frmETERM
        glbTermTran = True
        
    '    Call remNode
    '    Call TreeSetting
        
        Load frmETERM
        frmETERM.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Term:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_Tran_In_Click()
On Error GoTo Err_Term

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        glbDivTranInPlant = "N"
        
        Load frmETRANIN
        frmETRANIN.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Term:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_Tran_WFC_Div_Click()
On Error GoTo Err_Term

    If glbWFC And glbUserID = "999999999" Then
        MsgBox "The 999999999 account can not be used for Transfer Out.  Please log in under your employee number, and try again.", vbInformation + vbOKOnly, "Termination Not Allowed"
        Exit Sub
    End If

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        glbDivTranInPlant = "Y"
        
        Unload frmETERM
        glbTermTran = False
        Load frmETERM
        frmETERM.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Term:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_Tran_Out_Click()
On Error GoTo Err_Term

    If glbWFC And glbUserID = "999999999" Then
        MsgBox "The 999999999 account can not be used for Transfer Out.  Please log in under your employee number, and try again.", vbInformation + vbOKOnly, "Termination Not Allowed"
        Exit Sub
    End If

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        glbDivTranInPlant = "N"
        
        Unload frmETERM
        glbTermTran = False
        Load frmETERM
        frmETERM.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Term:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mmnu_VacEnt_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSVacEnt
        frmSVacEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_SickEnt_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSickEnt
        frmSickEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_VacEntDaily_Click()
        Screen.MousePointer = HOURGLASS
        Load frmSVacEntDaily
        frmSVacEntDaily.ZOrder 0
        Screen.MousePointer = DEFAULT
End Sub

Private Sub mmnu_HoursVacEntMst_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSHrsVacEnt
        frmSHrsVacEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_VacEarnedHours_Click()
    'If gSec_Upd_Entitlements Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmVacEarnedCalc
        frmVacEarnedCalc.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_VacPayPercentage_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSVacPayPrct
        frmSVacPayPrct.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_AnnVacEnt_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmSVacAnnEnt
        frmSVacAnnEnt.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_CurrentAccrYearEnd_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then
        Screen.MousePointer = HOURGLASS
        Load frmUCurAccYEnd
        frmUCurAccYEnd.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Benefit_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("BNCD") Then
        Screen.MousePointer = HOURGLASS
        Call Get_Code_Linamar("BNCD", "Benefit Codes", "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_EXTE_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        Unload frmTLAY
        glbTLAY = "Extending"
        Load frmTLAY
        frmTLAY.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_File_Door_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_DoorAccess Then
        Screen.MousePointer = HOURGLASS
        Unload frmLDoors
        Load frmLDoors
        frmLDoors.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_File_DoorName_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_DoorName Then
        Screen.MousePointer = HOURGLASS
        Load frmSDoorsName
        frmSDoorsName.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_File_EmailLog_Click()
    On Error GoTo ErrorHandler
        
    'Ticket #24629 - SMTP Log is now maintained in a table. So show the log in a report.
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCommonDlg.WindowShowPrintSetupBtn = True
    
    Me.vbxCommonDlg.ReportFileName = glbIHRREPORTS & "rzSMTPLog.rpt"
    Me.vbxCommonDlg.Connect = RptODBC_SQL
    Me.vbxCommonDlg.WindowTitle = "SMTP Log Report"
    Me.vbxCommonDlg.Destination = 0
    MDIMain.Timer1.Enabled = False
    Me.vbxCommonDlg.Action = 1
    vbxCommonDlg.Reset
    
    'Ticket #24629 - because of the above change, the following has been commented out.
'    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "SMTPLOG.TXT") = "" Then
'        MsgBox "No SMTP Log found.  This log is created automatically when an email is sent during termination.", vbExclamation + vbOKOnly, "No Log Found"
'        Exit Sub
'    End If
'    'Shell "START """ & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "SMTPLOG.TXT", vbNormalFocus
'    Dim dTaskID As Double
'    'Ticket #18175
'    dTaskID = Shell("Notepad """ & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "SMTPLOG.TXT""", vbNormalFocus)
'
'    'Shell App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "SMTPLOG.TXT", vbNormalFocus
    Exit Sub
    
ErrorHandler:
    'Ticket #24629 - because of the above change, the following has been commented out.
'    If Err.Number = 53 Then
'        MsgBox "Can't find START.EXE in path.  This is a core Windows component, usually found in \WINDOWS\COMMAND, or \WINNT\SYSTEM32.  Please ensure this file is correctly installed, and try again.", vbCritical + vbOKOnly, "START Not Found"
'        Exit Sub
'    Else
        MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
        Exit Sub
'    End If
End Sub

Private Sub mnu_File_EmailSetup_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Security Or gSec_Inq_Quick_ESS Then
        On Error Resume Next
        Load frmEMAIL
        frmEMAIL.ZOrder 0
        On Error GoTo 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_File_QuickSetupESS_Click()
    'If gSec_Inq_Security Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Quick_ESS Then
        Screen.MousePointer = HOURGLASS
        Load frmQuickESS
        frmQuickESS.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_WorkScheduleRule_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSWorkSchRule
    frmSWorkSchRule.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mnu_DashboardSetup_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSDashboardRule
    frmSDashboardRule.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mnu_OnCallHours_Click()
    Screen.MousePointer = HOURGLASS
    Load frmSOnCallHrs
    frmSOnCallHrs.ZOrder 0
    Screen.MousePointer = DEFAULT
End Sub

Private Sub mnu_Home_Line_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("HMLN") Then
        Screen.MousePointer = HOURGLASS
        Call Get_HOME("HMLN", "Home Line Codes", True, "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_Home_Operation_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("HMOP") Then
        Screen.MousePointer = HOURGLASS
        Call Get_HOME("HMOP", "Home Operation Number Codes", True, "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_Home_Shift_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("HMSF") Then
        Screen.MousePointer = HOURGLASS
        Call Get_HOME("HMSF", "Home Shift Codes", True, "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_Home_Work_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("HMWC") Then
        Screen.MousePointer = HOURGLASS
        Call Get_HOME("HMWC", "Home Work Center Codes", True, "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_HRONLY_Click()
    frmImportDb.Show 1
End Sub

Private Sub mnu_M_Update_Click()
    Call Me.ActiveForm.cmdModify_Click
End Sub

Private Sub mnu_NewEmployee_Click()
    If glbtermopen Then
        Call mmnu_Active_Click
    End If
    Call clkNew("NewEmployee")
End Sub

Private Sub mnu_Operation_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("EDSE") Then
        Screen.MousePointer = HOURGLASS
        Call Get_Code_Linamar("EDSE", "Operation Codes", "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_Pos_PositionCtrl_Click()
Dim fglbEditMode%, xPos
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then
        If fglbEditMode% Then
            MsgBox "Changes pending - save or cancel first"
            Exit Sub
        End If
        xPos = Not Len(glbPos$) = 0
        Screen.MousePointer = HOURGLASS
        Load frmPosControl
        frmPosControl.ZOrder 0
        If xPos Then frmPosControl.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Pos_DivDeptLnk_Click()
Dim fglbEditMode%, xPos
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then
        'If fglbEditMode% Then
        '    MsgBox "Changes pending - save or cancel first"
        '    Exit Sub
        'End If
        'xPos = Not Len(glbPos$) = 0
        Screen.MousePointer = HOURGLASS
        Unload frmMPOSITIONS
        frmPosDivDeptLnk.Show 1
        'frmPosDivDeptLnk.ZOrder 0
        'If xPos Then frmPosDivDeptLnk.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Pos_Budget_Click()
    Dim fglbEditMode%, xPos
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then 'And gSec_Inq_WHSCC_BUDPOS% Then
        'If Not frmMPOSITIONS.Data1.Recordset.EOF And Not frmMPOSITIONS.Data1.Recordset.BOF Then
            If fglbEditMode% Then
                MsgBox "Changes pending - save or cancel first"
                Exit Sub
            End If
            xPos = Not Len(glbPos$) = 0
            'glbPos$ = frmMPOSITIONS.txtPosition
            'glbPosDesc$ = frmMPOSITIONS.txtPosDescr
            'Unload frmMPOSITIONS
            Screen.MousePointer = HOURGLASS
            If glbWFC Then 'Ticket #25911 Franks 10/07/2014
                Load frmPosBudgetWFC
                frmPosBudgetWFC.ZOrder 0
                If xPos Then frmPosBudgetWFC.ZOrder 0
            Else
                Load frmPosBudget
                frmPosBudget.ZOrder 0
                If xPos Then frmPosBudget.ZOrder 0
            End If
            Screen.MousePointer = DEFAULT
        'Else
        '    MsgBox "No positions to select from"
        'End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Pos_Duties_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then '
        frmPosDuties.Show
        frmPosDuties.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Pos_Resp_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then '
        frmPosResp.Show
        frmPosResp.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Pos_AppProc_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then '
        frmPosAppProc.Show
        frmPosAppProc.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_Pos_Grid_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then '
        frmMPosGrid.Show
        frmMPosGrid.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Pos_Course_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then
       Load frmPosCourse
       frmPosCourse.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Shift_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("EDRG") Then
        Screen.MousePointer = HOURGLASS
        Call Get_Code_Linamar("SHFT", "Shift Codes", "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If

End Sub
Private Sub mnu_ProductLine_Click()
On Error GoTo ERR_EXIT
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("EDRG") Then
        Screen.MousePointer = HOURGLASS
        Call Get_Code_Linamar("EDRG", "Product Line Codes", "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_Skill_code_Click()
On Error GoTo ERR_EXIT

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Master_Table("EDSK") Then
        Screen.MousePointer = HOURGLASS
        Call Get_HOME("EDSK", "Skill Codes", True, "")
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
    
ERR_EXIT:
    If Err.Number = 5 Then
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
End Sub

Private Sub mnu_ProductLine_Operation_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Productline_Operation Then
        frmProductLineOperation.Show
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_REAT_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        Unload frmTLAY
        glbTLAY = "Re-activate"
        Load frmTLAY
        frmTLAY.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Term_Rehire_Click()
On Error GoTo Err_Rehire
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        Load frmEREHIRE
        frmEREHIRE.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Rehire:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_Term_DeathProcess_Click()
On Error GoTo Err_Rehire

        Screen.MousePointer = HOURGLASS
        UnloadFrms
        Screen.MousePointer = HOURGLASS
        'glbFrmCaption$ = "Death of a Retiree/Spouse"
        'Ticket #23491 Franks 04/02/2013
        glbFrmCaption$ = "Death of an Employee/Spouse"
        Unload frmERetirement
        Load frmERetirement
        frmERetirement.ZOrder 0
        Screen.MousePointer = DEFAULT

Exit Sub
Err_Rehire:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub
Private Sub mmnu_Term_RetWorking_Click()
On Error GoTo Err_Rehire

        Screen.MousePointer = HOURGLASS
        UnloadFrms
        Screen.MousePointer = HOURGLASS
        glbFrmCaption$ = "Retired – Working Process"
        Unload frmERetirement
        Load frmERetirement
        frmERetirement.ZOrder 0
        Screen.MousePointer = DEFAULT

Exit Sub
Err_Rehire:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_Term_RetRetiree_Click()
On Error GoTo Err_Rehire

        Screen.MousePointer = HOURGLASS
        UnloadFrms
        Screen.MousePointer = HOURGLASS
        glbFrmCaption$ = "Working Retiree Retirement Process"
        Unload frmERetirement
        Load frmERetirement
        frmERetirement.ZOrder 0
        Screen.MousePointer = DEFAULT

Exit Sub
Err_Rehire:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_Term_Retirement_Click()
On Error GoTo Err_Rehire

        Screen.MousePointer = HOURGLASS
        UnloadFrms
        Screen.MousePointer = HOURGLASS
        glbFrmCaption$ = "Retirement Process"
        Unload frmERetirement
        Load frmERetirement
        frmERetirement.ZOrder 0
        Screen.MousePointer = DEFAULT

Exit Sub
Err_Rehire:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub
Private Sub mmnu_Term_Rehire_Click()
On Error GoTo Err_Rehire
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = "        "
        UnloadFrms
        Screen.MousePointer = HOURGLASS
        Load frmEREHIRE
        frmEREHIRE.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

Exit Sub
Err_Rehire:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mmnu_VTile_Click()
    MDIMain.Arrange 2
End Sub

Private Sub mmnu_Windows_Click()
    MDIMain.panHelp(0).Caption = "Select a menu item"
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(3).Caption = ""
End Sub

Private Sub mnu_About_Click()
    'MenuAbout
    frmAbout.Show 1
End Sub

Private Sub mnu_File_ChgPass_Click()
'Dim Msg$, Def$, Response$, Title$, strNPass1$

    frmSPassCh.fdFrameName = "fraVerify"
    Load frmSPassCh
    frmSPassCh.Show
    'Msg$ = "Type new password."
    'Title$ = "Change Password"
    'Response$ = InputBox$(Msg$, Title$, Def$)
    'If Len(Response$) > 6 Then
    '    MsgBox "Password must be 6 characters or under.", vbExclamation + vbOKOnly, "Password Change Cancelled"
    '    Exit Sub
    'End If
    'If Len(Response$) > 0 Then
    '    strNPass1$ = Response$
    '    Msg$ = "Type new password again to verify."
    '    Title$ = "Verify Password Change"
    '    Response$ = InputBox$(Msg$, Title$, Def$)
    '    If Len(Response$) > 0 Then
    '        If Response$ <> strNPass1$ Then
    '            MsgBox "Password verification failed"
    '            Exit Sub
    '        Else
    '            glbPassword$ = strNPass1$
    '            Call modUpdPass(glbUserID, strNPass1$)
    '        End If
    '    End If
    'End If
End Sub

Private Sub mnu_File_Secure_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Security Then
        Screen.MousePointer = HOURGLASS
        Load frmSECURE
        frmSECURE.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_Information_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master Then
    If glbWFC Then 'Ticket #25911 Franks 10/06/2014
        Load frmMPOSITIONSWFC
        frmMPOSITIONSWFC.ZOrder 0
    Else
        Load frmMPOSITIONS
        frmMPOSITIONS.ZOrder 0
    End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_JobMaster_Click()
        Load frmMJobMasterMain
        frmMJobMasterMain.ZOrder 0
End Sub

Private Sub mnu_Pos_Eval_Click()
Dim fglbEditMode%, xPos

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Job_Master And gSec_Inq_Job_Eval Then
        'If Not frmMPOSITIONS.Data1.Recordset.EOF And Not frmMPOSITIONS.Data1.Recordset.BOF Then
            If fglbEditMode Then
                MsgBox "Changes pending - save or cancel first"
                Exit Sub
            End If
            xPos = Not Len(glbPos$) = 0
    '        glbPos$ = frmMPOSITIONS.txtPosition
    '        glbPosDesc$ = frmMPOSITIONS.txtPosDescr
            'Unload frmMPOSITIONS
            Screen.MousePointer = HOURGLASS
            Load frmPosEval
            frmPosEval.ZOrder 0
            If xPos Then frmPosEval.ZOrder 0
        'Else
        '    MsgBox "No positions to select from"
        'End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_F_PrintSetup_Click()
    vbxCommon.Action = 5
End Sub

Private Sub mnu_Pension_Click()
Dim xFile, xPath
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        xFile = App.Path
        xFile = xFile & IIf(Right(xFile, 1) = "\", "", "\")
        xPath = xFile & IIf(Right(xFile, 1) = "\", "", "\")
        
        xFile = xPath & "IHRPension.exe"
        If Dir(xFile) = "" Then
            MsgBox xFile & " not found"
        Else
            Call Shell(xFile & " " & glbUserID & "," & glbTxtPassword)
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnu_sysinfo_Click()
    Call SysInfo
End Sub

Private Sub mnu_Term_Inquiry_Click()
On Error GoTo Err_TermInq
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        glbtermopen = True
        glbTermTran = True
        Call GET_EMP
        If glbTERM_ID = 0 Then
            glbtermopen = False
        Else
            glbOnTop = ""
            UnloadFrms
            Call remNode
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT
Exit Sub

Err_TermInq:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mnu_Tran_Inquiry_Click()
On Error GoTo Err_TermInq

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        glbtermopen = True
        glbTermTran = False
        glbOnTop = "FRMETRANIN"
        
        Call GET_EMP
        
        If glbTran_ID = 0 Then
            glbtermopen = False
        Else
            glbOnTop = ""
            UnloadFrms
            Call remNode
        End If
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Screen.MousePointer = DEFAULT

Exit Sub
Err_TermInq:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub mnu_Tlay_Click()
On Error GoTo Err_Tlay

    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        Load frmETLAY
        frmETLAY.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
Exit Sub
Err_Tlay:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "MENU", "SELECT")
    If gintRollBack% = False Then Resume Next Else Unload Me
End Sub

Private Sub mnuUnionSickBank_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_WHSCC_USB Then
        Screen.MousePointer = HOURGLASS
        Load frmUSB
        frmUSB.ZOrder 0
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnuTerminationCauseLink_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_Company Then
        Screen.MousePointer = HOURGLASS
        frmEHSCauseLinks.LinkItem = "TERMCAUSE"
        frmEHSCauseLinks.Show 1
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mnuWorkFlowMaster_Click()
        Screen.MousePointer = HOURGLASS
        frmWorkFlowMaster.Show 1
        Screen.MousePointer = DEFAULT
End Sub

Private Sub submnu_Plan_Data_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Upd_EmploymentEQT Then
        Load frmPlanData
        frmPlanData.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub submnu_Rep_EmpEquityVitalAire_Click()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim snapFindPlan As New ADODB.Recordset

    'If gSec_Inq_EmploymentEQT Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayEQT Then
        Screen.MousePointer = HOURGLASS
    
        SQLQ = "SELECT * FROM HRPARCOP"
        snapFindPlan.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If snapFindPlan.EOF And snapFindPlan.BOF Then
            Msg = "No Plan Number descriptions found." & Chr(10)
            MsgBox Msg
            Screen.MousePointer = DEFAULT
            Exit Sub
        Else
            Load frmREmpEquity
            frmREmpEquity.ZOrder 0
            frmREmpEquity.Caption = "Employment Equity Report"
        End If
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If

End Sub

Private Sub submnu_Rep_ComplWork_Click()

Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim snapFindPlan As New ADODB.Recordset

    'If gSec_Inq_EmploymentEQT Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayEQT Then
        Screen.MousePointer = HOURGLASS
    
        SQLQ = "SELECT * FROM HRPARCOP"
        snapFindPlan.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If snapFindPlan.EOF And snapFindPlan.BOF Then
            Msg = "No Plan Number descriptions found." & Chr(10)
            MsgBox Msg
            Screen.MousePointer = DEFAULT
            Exit Sub
        Else
            If frmWorkForce.Visible Then
                frmWorkForce.Caption = "Completed Workforce Surveys"
                frmWorkForce.chkShowEmp.Visible = False
            Else
                Unload frmWorkForce
                Load frmWorkForce
                frmWorkForce.ZOrder 0
                frmWorkForce.Caption = "Completed Workforce Surveys"
            End If
        End If
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub submnu_Rep_EmplStatus_Click()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim snapFindPlan As New ADODB.Recordset

    'If gSec_Inq_EmploymentEQT Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayEQT Then
        Screen.MousePointer = HOURGLASS
        
        SQLQ = "SELECT * FROM HRPARCOP"
        snapFindPlan.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If snapFindPlan.EOF And snapFindPlan.BOF Then
            Msg = "No Plan Number descriptions found." & Chr(10)
            MsgBox Msg
            Screen.MousePointer = DEFAULT
            Exit Sub
        Else
            If frmWorkForce.Visible Then
                frmWorkForce.Caption = "Employment Status Analysis"
                frmWorkForce.chkShowEmp.Visible = True
            Else
                Unload frmWorkForce
                Load frmWorkForce
                frmWorkForce.ZOrder 0
                frmWorkForce.Caption = "Employment Status Analysis"
            End If
        End If
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub submnu_Rep_OccupGroup_Click()
    ' danielk - 12/31/2002 -  removed all the code that was here, someone copied submnu_Rep_Work_Click and
    '                         didn't change anything!  only thing that made it work was that there was
    '                         another code bug that made frmREEO pop up w/o clicking on it.
    Unload frmREEO
    glbFormCaption = "EEO Reports"
    frmREEO.Show
End Sub
Private Sub submnu_PurgeTermEEO_Click()
    Unload frmREEO
    glbFormCaption = "Purge Applicants EEO Records"
    frmREEO.Show
End Sub

Private Sub submnu_Rep_Work_Click()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim snapFindPlan As New ADODB.Recordset

    'If gSec_Inq_EmploymentEQT Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_PayEQT Then
        Screen.MousePointer = HOURGLASS
        
        SQLQ = "SELECT * FROM HRPARCOP"
        snapFindPlan.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If snapFindPlan.EOF And snapFindPlan.BOF Then
            Msg = "No Plan Number descriptions found." & Chr(10)
            Msg = Msg & "You will require authority to add one to continue"
            MsgBox Msg
            Screen.MousePointer = DEFAULT
            Exit Sub
        Else
            If frmWorkForce.Visible Then
                frmWorkForce.Caption = "Employer Workforce Survey"
                frmWorkForce.chkShowEmp.Visible = False
            Else
                Unload frmWorkForce
                Load frmWorkForce
                frmWorkForce.ZOrder 0
                frmWorkForce.Caption = "Employer Workforce Survey"
            End If
        End If
        Screen.MousePointer = DEFAULT
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub submnu_Survey_Data_Click()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String
Dim snapFindPlan As New ADODB.Recordset
    
    'If gSec_Upd_EmploymentEQT Then
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_EmploymentEQT Then
        Screen.MousePointer = HOURGLASS
        SQLQ = "SELECT * FROM HRPARCOP"
        snapFindPlan.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If snapFindPlan.EOF And snapFindPlan.BOF Then
            Msg = "No Plan Number descriptions found." & Chr(10)
            MsgBox Msg
            Screen.MousePointer = DEFAULT
            Exit Sub
        Else
            Load frmSurveyData
            frmSurveyData.ZOrder 0
        End If
        Screen.MousePointer = DEFAULT

    'Else
    '   MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub submnu_USData_Click()
    Load frmEEO
    frmEEO.ZOrder 0
End Sub

' danielk - 12/31/2002 - removed below code, made the form show up when you moved the pointer over the
'                        item.  this is a submenu, and the item that actually displays frmREEO is
'                        submnu_Rep_OccupGroup_Click()
'Private Sub submnu_USReports_Click()
'    Load frmREEO
'End Sub
' danielk - 12/31/2002 - end

Private Sub Timer1_Timer()
 
lblTime = Format(Now, "Short Date") & " - " & Format(Now, "Medium Time")

End Sub

Sub MuneSetup(TF)
    'mmnu_EE_comPlan.Visible = (glbCompSerial = "S/N - 2291W" Or glbCompSerial = "S/N - 2325W") And TF
    ''mmnu_R_ComPlan.Visible = (glbCompSerial = "S/N - 2291W" Or glbCompSerial = "S/N - 2325W") And TF  'Removing Reports from Menu Bar
    
    'mmnu_Home_Master.Visible = glbLinamar
    ''mmnu_R_DoorAccess.Visible = glbLinamar
    'mnu_File_Door.Visible = glbLinamar
    'mnu_File_DoorName.Visible = glbLinamar
End Sub

Private Sub MDIForm_Resize()
    If Me.Height - panMain.Height - MainToolBar.Height - 900 < 0 Then
        tvwTree.Height = 0
        lstView.Height = 0
    Else
        tvwTree.Height = Me.Height - panMain.Height - MainToolBar.Height - 900
        lstView.Height = Me.Height - panMain.Height - MainToolBar.Height
        lstPanel.Width = Me.Width - panTree.Width
        lstView.Width = lstPanel.Width
    End If
    fraMove.Height = panTree.Height
End Sub

Private Sub TreeSetting()
Dim xparent As String
Dim xTmpFlag As Boolean 'Ticket #24155 Franks 07/29/2013

On Error GoTo err_TreeSetting

tvwTree.Nodes.Add , , "root", "info:HR", "infohr"
xparent = "root"
addNode xparent, "Employee", "key1", "applicants"
If glbWFC Then 'Ticket #25911 Franks 10/07/2014
    addNode xparent, "Job/Position", "key2", "positions"
Else
    addNode xparent, "Position", "key2", "positions"
End If
If glbLinamar Then
    addNode xparent, lStr("Division Health & Safety"), "keyHSDiv", "positions"
End If
addNode xparent, "Reports", "key3", "reports"
If Not glbtermopen Then
    addNode xparent, "Mass Updates", "key4", "mass"
    If glbWFC Then 'Ticket #29013 Franks 08/23/2016
        If glbWFC_IncentivePlanFlag Then
            addNode xparent, "Incentive Plan", "keyIP", "applicants"
        End If
    End If
    addNode xparent, "Setup", "key5", "setup"
End If
If File_Exist("IHREI.exe") Then 'Ticket #13142
    addNode xparent, "Import/Export", "keypEI", "payweb"
End If
If Not glbVadim Or glbPayWeb Then
    addNode xparent, "Payweb", "keypw6", "payweb"
End If
If glbVadim Then
    addNode xparent, "Vadim", "keyvd7", "payweb"
End If
If glbAdv Or glbWFCFullRights Then
    addNode xparent, "Advanced Tracker", "keyat8", "payweb"
    'addNode xparent, "Advanced Tracker2", "keyat8", "payweb"
End If
''geo added begin
If glbGP Then
    addNode xparent, "Great Plains", "keyat9", "payweb" '
End If
'
If glbMediPay Then
    addNode xparent, "MediPay", "keyat10", "payweb" '
End If
''geo added end

If glbWFC Or glbCompSerial = "S/N - 9999W" Then 'Ticket #24184 Franks 09/11/2013
    'Ticket #25522 Franks 05/23/2014 - add 9999
    addNode xparent, "HRSoft", "keyat12", "payweb"
End If

If glbCompSerial = "S/N - 2379W" Then  'Ticket #26912 Franks 06/22/2015
    addNode xparent, "System 24/7", "keyat15", "payweb"
End If

'Simona - Leeds Grenville CAS ticket #14890
If glbCwis Then
    addNode xparent, "CWIS", "keyat11", "payweb"
End If

'If glbWFC Then
'    addNode xparent, "Bonus System", "keyat11", "payweb"
'End If

xparent = "key1"
    If glbWFC Or glbCompSerial = "S/N - 9999W" Then ''Ticket #25522 Franks 05/23/2014 - add 9999
        If Not glbtermopen Then
            addNode xparent, "Find Candidate ", "KeyFindCandi", "find"
        End If
    End If
    If Not glbtermopen Then
        addNode xparent, "Find Active Employee", "Ke", "find"
    End If
    If glbTermTran Or Not glbtermopen Then
        If gSec_Inq_Terminations Then   'Ticket #16189
            addNode xparent, "Find Terminated Employee", "key6", "find"
        End If
    End If

    addNode xparent, "Basic Information", "key7", "applicants"
    addNode xparent, "Work History/Compensation", "key8", "applicants"
    addNode xparent, "Attendance/Entitlements", "key9", "applicants"
    addNode xparent, "Education/Skills", "key10", "applicants"
    'If Not glbtermopen Then    'Ticket #18668
        addNode xparent, lStr("Follow-ups"), "key11", "applicants"    'Ticket #15088
    'End If
    addNode xparent, "Health & Safety", "key12", "applicants"
    If gSec_Inq_Counselling Then    'Ticket #16189
        If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
            addNode xparent, "Assets", "key13", "applicants"
        Else
            addNode xparent, lStr("Counseling"), "key13", "applicants"
        End If
    End If
    If gSec_Inq_Comments Then   'Ticket #16189
        addNode xparent, lStr("Comments"), "key14", "applicants"
    End If
    If glbUS Then
        addNode xparent, "COBRA", "key14a", "applicants"
    End If
    If glbLinamar Then
        addNode xparent, "Terminations", "key15", "applicants"
    Else
        addNode xparent, "Leaves and Terminations", "key15", "applicants"
    End If
    If Not glbtermopen = False Then
        addNode xparent, "Switch To Active Employees", "tkey15", "applicants"
    End If
xparent = "key7"
    If gSec_Inq_Basic Then  'Ticket #16189
        addNode xparent, "Employee Demographics", "Key71", "applicants"
        addNode xparent, "Status/Dates", "key72", "applicants"
    End If
    If gSec_Inq_EmergContacts Then  'Ticket #17503
        addNode xparent, "Emergency Contacts", "key73", "applicants"
    End If
    If gSec_Inq_Dependents Then     'Ticket #16189
        addNode xparent, "Dependents", "key74", "applicants"
    End If
    If gSec_Inq_Banking Then    'Ticket #16189
        addNode xparent, "Banking Information", "key75", "applicants"
    End If
    If gSec_Inq_GLDist Then 'Ticket #16189
        addNode xparent, lStr("G/L") & " Distribution", "KeyGLDist", "applicants" ' added by Bryan 23/Feb/06 Ticket#10308
    End If
    If gSec_Inq_OtherInformation Then   'Ticket #16189
        addNode xparent, "Other Information", "key76", "applicants"
    End If
    If gSec_Inq_EMP_HISTORY Then    'Ticket #16189
        addNode xparent, "Employee History", "key77", "applicants"
    End If
    If gSec_Inq_EMP_FLAGS Then  'Ticket #16189
        addNode xparent, "Employee Flags", "key78", "applicants"
    End If
    If gSec_Inq_ADP_Data Then   'Ticket #16189
        addNode xparent, "Employee ADP Data", "key79", "applicants"
    End If
    If gSec_Inq_AddPayrollIDData Then   'Ticket #25015 - Macaulay
        addNode xparent, lStr("Additional Payroll ID Data"), "key88", "applicants"
    End If
    
xparent = "key8"
    If gSec_Inq_Position Then   'Ticket #16189
        addNode xparent, "Position", "Key81", "applicants"
    End If
    If gSec_Inq_Salary Then 'Ticket #16189
        addNode xparent, "Salary", "key82", "applicants"
    End If
    
    If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation - Ticket #10844
        If gSec_Upd_Performance Then    'Ticket #16189
            addNode xparent, "Staff Profile", "PerfReview", "applicants"
        End If
    Else
        If gSec_Inq_Performance Then    'Ticket #16189
            addNode xparent, lStr("Performance"), "key83", "applicants"
        End If
    End If
    If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation - Ticket #10844
        If gSec_Inq_Temp_Cross_Training Then    'Ticket #16189
            addNode xparent, "Temporary/Cross Training Assignment", "Key81a", "applicants"
        End If
    End If
    If gSec_Inq_Earnings Then   'Ticket #16189
        addNode xparent, "Other Earnings", "key84", "applicants"
    End If
    If gSec_Inq_Benefits Then   'Ticket #16189
        addNode xparent, "Benefits/Beneficiaries", "key85", "applicants"
    End If
    If gSec_Inq_Other_Entitlements Then 'Ticket #16189
        addNode xparent, "Dollar Entitlements", "key86", "applicants"
    End If
    If glbCompSerial = "S/N - 2242W" Then   'C.C.A.C. London & Middlesex - Ticket #6718
        If gSec_Inq_SalDist Then    'Ticket #16189
            addNode xparent, "Salary Distribution", "eesaldist", "applicants"
        End If
    End If
    
    If gSec_Inq_PayrollTrans Then   'Ticket #16189
        addNode xparent, "Payroll Transactions", "keyPayTrans", "applicants" 'Ticket #13035
    End If
    
    addNode xparent, "Previous Work History", "key87", "applicants"
    xparent = "key87"
    If gSec_Inq_Position Then   'Ticket #16189
        addNode xparent, "Position", "key871", "applicants"
    End If
    If gSec_Inq_Salary Then 'Ticket #16189
        addNode xparent, "Salary", "key872", "applicants"
    End If
    
    If glbCompSerial = "S/N - 2279W" Then 'Friesens Corporation - Ticket #10844
        If gSec_Upd_Performance Then    'Ticket #16189
            addNode xparent, "Staff Profile", "PerfReviewH", "applicants"
        End If
    Else
        If gSec_Inq_Performance Then    'Ticket #16189
            addNode xparent, lStr("Performance"), "key873", "applicants"
        End If
    End If

    If glbCompSerial = "S/N - 2382W" Then 'Samuel Ticket #20052 Franks 07/25/2011
        xparent = "key8"
        If gSec_Inq_Profit_Sharing Then
            addNode xparent, "Profit Sharing", "key875", "applicants"
        End If
    End If
    
xparent = "key9"
    If glbWHSCC Then
        If gSec_Inq_WHSCC_ASL Then  'Ticket #16189
            addNode xparent, "Advance Sick Leave", "key90a", "applicants"
        End If
    End If
    If gSec_Inq_Attendance Then 'Ticket #16189
        addNode xparent, "Attendance", "Key91", "applicants"
    End If
    If (Not glbtermopen) Then   'Ticket #22117
        If gSec_Inq_Attendance_History Then 'Ticket #16189
            addNode xparent, "Attendance History", "key92", "applicants"
        End If
    End If
    If glbWFC Or glbCompSerial = "S/N - 2418W" Then
        If gSec_Inq_Entitlements Then   'Ticket #16189
            addNode xparent, "Vacation Entitlements", "key93", "applicants"
            addNode xparent, "Vacation Overview", "key94", "applicants"
        End If
    Else
        If glbCompSerial <> "S/N - 2380W" Then 'VitalAire Ticket #13979
            If gSec_Inq_Entitlements Then   'Ticket #16189
                addNode xparent, "Vacation and Sick Entitlements", "key93", "applicants"
            End If
        End If
        If gSec_Inq_Entitlements Then   'Ticket #16189
            addNode xparent, "Vacation and Sick Overview", "key94", "applicants"
        End If
    End If
    If gSec_Inq_Hrly_Entitlements Then  'Ticket #16189
        addNode xparent, "Hourly Entitlements", "key95", "applicants"
    End If
    If gSec_Inq_Ovt_Overview Then   'Ticket #16189
        If glbCompSerial = "S/N - 2425W" Then 'Four Villages (Ticket #19998)
            addNode xparent, "Extra Time Bank Overview", "key96", "applicants"
        Else
            addNode xparent, "Overtime Bank Overview", "key96", "applicants"
        End If
    End If
    If Not glbtermopen Then 'Ticket #22743
        If gSec_Inq_Work_Schedule Then
            addNode xparent, "Work Schedule", "key97", "applicants"
        End If
    End If
    
xparent = "key10"
    If gSec_Inq_Associations Then   'Ticket #16189
        addNode xparent, lStr("Associations"), "Key101", "applicants"
    End If
    If gSec_Inq_Education_Seminars Then 'Ticket #16189
        addNode xparent, "Continuing Education", "key102", "applicants"
    End If
    If gSec_Inq_Formal_Education Then   'Ticket #16189
        addNode xparent, "Formal Education", "key103", "applicants"
    End If
    If gSec_Inq_EMP_LANG Then   'Ticket #16189
        addNode xparent, "Languages", "key106", "applicants"
    End If
    If gSec_Inq_Skills Then 'Ticket #16189
        addNode xparent, "Skills", "key104", "applicants"
    End If
    If glbLinamar Then
        If gSec_Inq_LinamarSkills Then  'Ticket #16189
            addNode xparent, "Skills for Production ", "key105", "applicants"
        End If
    End If
    If gSec_Inq_SUCCESSION Then 'Ticket #16189
        addNode xparent, "Succession Planning", "key107", "applicants"
    End If
    '7.9 - Enhancement - Add this option for all - the City of Chatham-Kent logic only
    'Friesens - Ticket #16189
    'City of Chatham-Kent - Ticket #16794
    'If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then  'Friesens Corporation - Ticket #10844
        If (Not glbtermopen) And gSec_Inq_Training_List Then
            'Ticket #20447 - Jerry asked to change to Training Plan for everyone except Friesens and
            'Chatham-Kent but Chatham-Kent are not using 7.9
            If glbCompSerial = "S/N - 2279W" Or glbCompSerial = "S/N - 2188W" Then
                addNode xparent, "Training List", "Key109", "applicants"
            Else
                addNode xparent, "Training Plan", "Key109", "applicants"
            End If
        End If
    'End If
    
    If gSec_Inq_UserDefineTbl Then  'Ticket #16189
        addNode xparent, lStr("User Defined Table"), "Key108", "applicants"
    End If

xparent = "key11"
    If Not glbtermopen Then
        If gSec_Inq_Follow_Ups Then 'Ticket #16189
            addNode xparent, lStr("Follow-ups Maintenance"), "Key111", "applicants"
            addNode xparent, lStr("Follow-ups Overview"), "key112", "applicants"
            If glbSQL Then
                If glbWFC Then
                    addNode xparent, ("Work Flow Overview"), "key113", "applicants"
                End If
            End If
        End If
    Else
        'Ticket #18668
        If gSec_Inq_Follow_Ups Then 'Ticket #16189
            addNode xparent, lStr("Follow-ups Maintenance"), "Key111", "applicants"
        End If
    End If
xparent = "key12"
    If gSec_Inq_Health_Safety Then  'Ticket #16189
        addNode xparent, "Incident Reporting", "Key121", "applicants"
        addNode xparent, "Injury", "key122", "applicants"
        
        'Double security as you have to have Health & Safety to get to this form of WSIB Form 7
        If gSec_Inq_HSW7Injury And glbWSIBModule Then
            addNode xparent, "Injury WSIB Form 7", "key122a", "applicants"
        End If
        
        'Double security as you have to have Health & Safety to get to this form of WSIB Form 9
        If gSec_Inq_HSWF9 And glbWSIBModule Then
            'Ticket #21463
            addNode xparent, "WSIB Form 9", "key122b", "applicants"
        End If
        
        If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
            addNode xparent, "Reoccurrence", "key1221", "applicants"
        End If
    End If
    If gSec_Inq_HSRootCause Then    'Ticket #16189
        addNode xparent, "Root Causes", "key123", "applicants"
    End If
    If gSec_Inq_HSCorrectiveAct Then     'Ticket #16189
        addNode xparent, "Corrective Actions", "key124", "applicants"
    End If
    If gSec_Inq_HSClaimMed Then    'Ticket #16189
        addNode xparent, "Claim/Medical Information", "key125", "applicants"
    End If
    If gSec_Inq_HSContacts Then     'Ticket #16189
        addNode xparent, "Contacts", "key126", "applicants"
    End If
    If gSec_Inq_HSCost Then
        addNode xparent, "WSIB Cost Statements", "key127", "applicants"
        addNode xparent, "Company Associated Costs", "key130", "applicants"
    End If
    If gSec_Inq_Health_Safety Then
        If glbWFC Then
            addNode xparent, "Accident Cost Analysis", "key128", "applicants"
        End If
        addNode xparent, "Incident Documents", "key129", "applicants"
    End If

xparent = "key15"
    If Not glbtermopen Then
        If glbLinamar Then
            addNode xparent, "Temporary Lay Off", "key151-0", "applicants"
            If gSec_Inq_Terminations Then   'Ticket #16189
                addNode xparent, "Transfer Out", "key151-1", "applicants"
                addNode xparent, "Transfer In", "key151-2", "applicants"
            End If
        Else
            If glbWFC Or glbSamuel Then 'Ticket #20884 Franks 10/20/2011
                If gSec_Inq_Terminations Then   'Ticket #16189
                    addNode xparent, "Transfer Out", "key151-1", "applicants"
                    addNode xparent, "Transfer In", "key151-2", "applicants"
                    If glbWFC Then 'Ticket #25221 Franks 03/17/2014
                        addNode xparent, "Transfer Division within Plant", "key151-3", "applicants"
                    End If
                End If
            End If
            If gSec_Inq_EnterLeave Then 'Ticket #16189
                addNode xparent, "Enter a Leave", "key151", "applicants"
                addNode xparent, "LOA Date Change", "key152", "applicants"
                addNode xparent, "Re-Activate from a Leave", "key153", "applicants"
            End If
        End If
        If gSec_Inq_Terminations Then   'Ticket #16189
            addNode xparent, "Termination", "key154", "applicants"
        End If
    End If
    If gSec_Inq_Rehire Then   'Ticket #16189
        addNode xparent, "Rehire", "key155", "applicants"
    End If
    If glbWFC Then 'Ticket #18566
        If gSec_Inq_RetirementProc Then
            addNode xparent, "Retirement Process", "key156", "applicants"
        End If
        If Not glbtermopen Then
            If gSec_Inq_DeathProc Then
                'addNode xparent, "Death of a Retiree/Spouse", "key157", "applicants"
                'Ticket #23491 Franks 04/02/2013
                addNode xparent, "Death of an Employee/Spouse", "key157", "applicants"
            End If
        End If
    End If
xparent = "key151-0"
    If Not glbtermopen Then
        If glbLinamar Then
            If gSec_Inq_Terminations Then   'Ticket #16189
                addNode xparent, "Temporary Lay Off", "key151", "applicants"
                addNode xparent, "Extending", "key152", "applicants"
                addNode xparent, "Re-Activate", "key153", "applicants"
            End If
        End If
    End If
    
If glbLinamar Then
    xparent = "keyHSDiv"
    If gSec_Inq_Health_Safety Then  'Ticket #16189
        addNode xparent, "Incident Reporting", "keyHSD1", "positions"
        addNode xparent, "Injury", "keyHSD2", "positions"
        addNode xparent, "Root Causes", "keyHSD3", "positions"
        addNode xparent, "Corrective Actions", "keyHSD4", "positions"
        addNode xparent, "Claim/Medical Information", "keyHSD5", "positions"
        addNode xparent, "Contacts", "keyHSD6", "positions"
        addNode xparent, "Incident Documents", "keyHSD7", "positions"
    End If
End If

xparent = "key2"
    
    If glbWFC Then 'Ticket #25911 Franks 09/24/2014
        addNode xparent, "Find Job", "Key16JB", "find" '"positions"
        If gSec_Inq_Job_Master Then 'Ticket #16189
                addNode xparent, "Job Master", "Key16JC", "positions"
        End If
        addNode xparent, "Find Position", "Key16a", "find"
        If gSec_Inq_Job_Master Then 'Ticket #16189
                addNode xparent, "Position Master", "Key16", "positions"
        End If
    Else
        addNode xparent, "Find Position", "Key16a", "find"
        If gSec_Inq_Job_Master Then 'Ticket #16189
            addNode xparent, "Master", "Key16", "positions"
        End If
    End If
    
    If glbMultiGrid Then
        If gSec_Inq_Job_Master Then 'Ticket #16189
            addNode xparent, "Salary Grid Details", "keyPosGrid", "positions"
        End If
    End If
    If gSec_Inq_Job_Master And gSec_Inq_Job_Skills Then 'Ticket #16189
        addNode xparent, "Skills", "key17", "positions"
    End If
        
    If gSec_Inq_Job_Master And gSec_Inq_Job_Eval Then   'Ticket #16189
        addNode xparent, "Eval Factors", "key18", "positions"
    End If
    If gSec_Inq_ReqCourses Then 'Ticket #16189
        addNode xparent, "Required Courses", "key19", "positions"
    End If
    If gSec_Inq_BudgetedPos Then 'Ticket #16189
        addNode xparent, "Budgeted Positions", "key191", "positions" 'Added By Frank Jun 30,2003 Ticket 4352
    End If
    If gSec_Inq_Job_Classes Then    'Ticket #16189
        addNode xparent, "National Occupation Classification", "key19a", "positions"
    End If
    If gSec_Inq_Job_Master Then 'Ticket #16189
        addNode xparent, "Position Duties", "keyPosDuties", "positions"
        addNode xparent, "Position Requirements", "keyPosResp", "positions"
    End If
    If gSec_Inq_AppProcess Then 'Ticket #16189
        addNode xparent, "Application Process", "keyPosAppProc", "positions"
    End If
    If gSec_Inq_Job_Master Then 'Ticket #16189
        If glbOttawaCCAC Then addNode xparent, "CCAC Positions", "key192", "positions"
    End If
    If gSec_Inq_SalaryGrids Then    'Ticket #16189
        If glbWFC Then addNode xparent, "Salary Grids", "key19b", "positions"
    End If
    'Friesens Corporation - Ticket #16658
    If glbCompSerial = "S/N - 2279W" Then
        If gSec_Inq_Job_Master Then 'Ticket #16658
            addNode xparent, lStr("Division") & " and " & lStr("Department") & " Link", "key193", "positions"
        End If
    End If
    

xparent = "key3"
    If Not glbtermopen Then
        addNode xparent, "Attendance and Entitlements", "key20", "reports"
        addNode xparent, "Employee Information", "key21", "reports"
        If gSec_Rpt_Counselling Then
            If glbCompSerial = "S/N - 2376W" Then ' George added for Assembling of 1st Nations #9535
                addNode xparent, "Assets", "key22", "reports"
            Else
                addNode xparent, lStr("Counseling"), "key22", "reports"
            End If
        End If
        
        If Not glbtermopen And gsAttachment_DB Then
            If gSec_Rpt_DocumentType Then  'Ticket #27244
                addNode xparent, "Document Type", "key222", "reports"
            End If
        End If
        
        If glbLinamar Then
            addNode xparent, "Door Access", "key22-L1", "reports"
        End If
        addNode xparent, "Educations/Skills", "key23", "reports"
        If gSec_Rpt_Follow_Ups Then
            addNode xparent, lStr("Follow-ups"), "key24", "reports"
            addNode xparent, lStr("Follow-ups Email Log"), "key24a", "reports"
        End If
        'Friesens Corporation - Ticket #16189
        If glbCompSerial = "S/N - 2279W" Then
            addNode xparent, "Friesens Forms", "key80", "reports"
        End If
        
        'Ticket #24663 - Add new menu under reports - Forms
        addNode xparent, "Forms", "key3Forms", "reports"
        
        addNode xparent, "Health and Safety", "key25", "reports"
    End If
   
    If gSec_Rpt_Master_Job Then 'Ticket #16189
        If glbWFC Then 'Ticket #25911 Franks 01/28/2015
            addNode xparent, "Positions/Skills/Evaluation", "key27wfc", "reports"
        Else
            addNode xparent, "Positions/Skills/Evaluation", "key27", "reports"
        End If
    End If
    If Not glbtermopen Then
        If gSec_Rpt_Seniority Then  'Ticket #16189
            addNode xparent, "Seniority", "key28", "reports"
        End If
    End If
    
    
    addNode xparent, "Setup", "key29", "reports"
    If glbCompSerial = "S/N - 2369W" Then
        addNode xparent, "Shift Schedule", "tsShiftRpt", "reports" 'added by Bryan 31/Oct/05 Tiocket#9630
    End If
    If gSec_Rpt_Master_Termination Then
        addNode xparent, "Terminations", "key30", "reports"
    End If
    If Not glbtermopen Then
        addNode xparent, "Work History/Compensation", "key31", "reports"
        
        'Hemu - add a new node for Statistical reports
        '     - move the Population Statistic report here from Employee Info.
        addNode xparent, "Statistics", "key32a", "reports"
    End If
    
xparent = "key20"
    If Not glbtermopen Then
        If gSec_Rpt_Entitlements Then   'Ticket #16189
            addNode xparent, "Accrual", "accrualreport", "reports"
        End If
        If gSec_Rpt_Master_Attendance Then  'Ticket #16189
            addNode xparent, "Attendance", "key201", "reports"
        End If
        If gSec_Rpt_Master_Attendance Then  'Ticket #16189
            'Release 8.0 - Ticket #22682    - Jerry asked me to move to Reports menu under Attendance and Entitlements
            addNode xparent, "Attendance Audit Master", "key47e", "reports"
        End If
        If gSec_Rpt_Bonus_Attendance Then  'Ticket #16189
            addNode xparent, "Attendance Bonus Points", "key202", "reports"
        End If
        If gSec_Rpt_Calendar_Attendance Then  'Ticket #16189
            addNode xparent, "Attendance Calendar", "key203", "reports"
        End If
        If gSec_Rpt_Attendance_Hist Then 'Ticket #16189
            addNode xparent, "Attendance History", "key204", "reports"
        End If
        
        If glbCompSerial = "S/N - 2425W" Then   'Four Villages - 'Ticket #21873
            If gSec_Rpt_AttWrkSch_Descrepancy Then
                addNode xparent, "Attendance/Work Schedule Discrepancy", "Descrepancy", "reports"
            End If
        End If
        
        If gSec_Rpt_Compensatory_Time Then  'Ticket #16189
            If glbCompSerial = "S/N - 2425W" Then   'Four Villages - Ticket #19998
                addNode xparent, "Extra Time", "key205", "reports"
            Else
                addNode xparent, "Compensatory Time", "key205", "reports"
            End If
        End If
                
        'If gSec_Rpt_Cost_Of_Employment And gSec_Inq_Attendance_History Then 'Ticket #16189
        If gSec_Rpt_Costed_Attendance Then 'Ticket #16189
            addNode xparent, "Costed Attendance", "key206", "reports"
        End If
        
        'Ticket #29230 - Daily Vacation Entitlement
        '              - Only display this screen if Vacation Enttitlement Earned is Daily and Vacation Entitlement Outstanding Based On = 1
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            If glbCompEntVacDaily And glbCompEntVac$ = "D" And glbEntOutStanding$ = "1" Then
                'addNode xparent, "Daily Accrual Skipped Log", "key392c", "reports"
                addNode xparent, "Daily Accrual", "key392DAR", "reports"
            End If
        End If
        
        If gSec_Rpt_Emergency_Leave Then    'Ticket #16189
            addNode xparent, "Emergency Leave", "key207", "reports"
        End If
        If gSec_Rpt_Entitlements Then   'Ticket #16189
            addNode xparent, "Entitlements", "key208", "reports"
        End If
        
        'County of Wellington - 'Ticket #22034
        If glbCompSerial = "S/N - 2262W" Then
            If gSec_Rpt_EnviroServices Then
                addNode xparent, "Wellington Terrace Attendance", "EnviroServ", "reports"
            End If
        End If
        
        If gSec_Rpt_ESSReq_TransAudit Then  'Ticket #16189
            addNode xparent, "ESS Requests - Transaction Audit", "ESSReqTrnAud", "reports"
        End If
        
        If glbBrantCount Then
            addNode xparent, "Flex Bank", "key2226_1", "reports"
        End If
        
        If gsec_rpt_Future_Entitlement Then 'Ticket #16189
            addNode xparent, "Future Entitlements", "key209a", "reports"
        End If
        If gSec_Rpt_Master_HourEnt Then 'Ticket #16189
            addNode xparent, "Hourly Entitlements", "key209", "reports"
        End If
        If glbCompSerial = "S/N - 2192W" Then
            If gSec_Rpt_Master_Attendance Then  'Ticket #16189
                addNode xparent, "Journal Entry", "JournalEntry ", "reports" ' county essex
            End If
        End If
        If gSec_Rpt_Ovt_Bank Then   'Ticket #16189
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
                addNode xparent, "Extra Time Bank", "key205a", "reports"
            Else
                addNode xparent, "Overtime Bank", "key205a", "reports"
            End If
        End If
        If gSec_Rpt_Ovt_Lost_Hours Then 'Ticket #16189
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
                addNode xparent, "Extra Time Bank Lost Hours", "key205b", "reports"
            Else
                addNode xparent, "Overtime Bank Lost Hours", "key205b", "reports"
            End If
        End If
        If glbCompSerial = "S/N - 2192W" Then ' county essex
            If gSec_Rpt_Master_Attendance Then  'Ticket #16189
                addNode xparent, "Timesheet", "timesheet", "reports"
                addNode xparent, "Timesheet With Equipment Cost", "timesheetWCost", "reports"
            End If
        'End If
        Else
        'Ticket #28002 - Opening for all clients and adding Termination employee option as well
        'If glbCompSerial = "S/N - 2174W" Then ' Kawartha Haliburton CAS
            If gSec_Rpt_Master_Attendance Then  'Ticket #16189
                addNode xparent, "Timesheet", "timesheet", "reports"
            End If
        End If
        If glbCompSerial = "S/N - 2241W" Then ' Granite Club Ticket #15133
            If gSec_Rpt_Master_Attendance Then  'Ticket #16189
                addNode xparent, "Personal Day Report", "PersonalDayRpt", "reports"
            End If
        End If
        'If glbCompSerial = "S/N - 2257W" Then ' Hamilton C.C.A.S   - Jerry said it's for everyone
            If gSec_Rpt_Master_Attendance Then  'Ticket #16189
                addNode xparent, "Timesheet Status", "key210", "reports"
            End If
        'End If
        If gSec_Rpt_Work_Schedule Then
            addNode xparent, "Work Schedule", "key99", "reports"
        End If
    End If
    
    'Ticket #30498
    If Not glbtermopen Then
        'Ticket #29230 - Daily Vacation Entitlement
        '              - Only display this screen if Vacation Enttitlement Earned is Daily and Vacation Entitlement Outstanding Based On = 1
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            If glbCompEntVacDaily And glbCompEntVac$ = "D" And glbEntOutStanding$ = "1" Then
                xparent = "key392DAR"
                    addNode xparent, "Details File", "key392d", "reports"
                    addNode xparent, "Skipped Log", "key392c", "reports"
            End If
        End If
    End If

xparent = "key21"
    If Not glbtermopen Then
        If gSec_Rpt_Age Then    'Ticket #16189
            addNode xparent, "Birthday/Age", "key211", "reports"
        End If
        If glbOttawaCCAC Then
            If gSec_Rpt_Job_List Then   'Ticket #16189
                addNode xparent, "Category/Status", "key2343_1", "reports"
            End If
        End If
        If gSec_Rpt_Dependents Then 'Ticket #16189
            addNode xparent, "Dependents", "key212", "reports"
        End If
        If gSec_Rpt_EmailAddress Then   'Ticket #16189
            addNode xparent, "Email Address", "key213", "reports"
        End If
        If gSec_Rpt_Emergecy_Contacts Then  'Ticket #16189
            addNode xparent, "Emergency Contacts", "key215", "reports"
        End If
        If gSec_Rpt_Profiles Or gSec_Rpt_Comments Then  'Ticket #16189
            addNode xparent, lStr("Employee/Comments"), "key214", "reports"
        End If
        
        ''Release 8.0 - Ticket #22682
        If gSec_Rpt_Employee_Dates Then 'Ticket #16189
            addNode xparent, "Employee Dates", "key812", "reports"
        End If
        
        If gSec_Rpt_Employee_Flags Then 'Ticket #16189
            addNode xparent, "Employee Flags", "key220", "reports"
        End If
        If gSec_Rpt_GLDistribution Then '79 - Ticket #18668
            addNode xparent, lStr("G/L") & " Distribution", "key811", "reports"
        End If
        If gSec_Rpt_Employee_Labels Then    'Ticket #16189
            addNode xparent, "Employee Labels", "key216", "reports"
        End If
        If gSec_Rpt_Turnover Then   'Ticket #16189
            addNode xparent, "Employee Turnover", "key217", "reports"
        End If
        If gSec_Rpt_Profiles Then   'Ticket #16189
            addNode xparent, "Employee Profile", "key218", "reports"
        End If
        If gSec_Rpt_Job_List Then   'Ticket #16189
            addNode xparent, "Employee/Position", "key219", "reports"
        End If
        If gSec_Rpt_Employee_Hist Then    'Ticket #16189
            addNode xparent, "Employee History", "key2114", "reports"
        End If
        If gSec_Rpt_Home_Address And gSec_Show_ADDRESS Then 'Ticket #16189
            addNode xparent, "Home Address/Phone", "key2110", "reports"
        End If
        If gSec_Rpt_LOA Then   'Ticket #16189
            addNode xparent, "Leave of Absence", "key2141", "reports"
        End If
        
        ''Release 8.0 - Ticket #22682
        If gSec_Rpt_Length_Of_Service Then 'Ticket #16189
            addNode xparent, "Length of Service", "key813", "reports"
        End If
        
        If gSec_Rpt_POE Then 'Ticket #16189
            addNode xparent, "Plan of Establishment", "key2113", "reports"
        End If
        'addNode xparent, "Population Statistics", "key2111", "reports"
        If gSec_Rpt_SINSSN And gSec_Show_SIN_SSN Then    'Ticket #16189
            addNode xparent, "S.I.N./S.S.N.", "key2111a", "reports"
        End If
        If gSec_Rpt_Telephone_Extensions Then   'Ticket #16189
            addNode xparent, "Telephone Extension", "key2112", "reports"
        End If
    End If
   
xparent = "key23"
    If Not glbtermopen Then
        If gSec_Rpt_Associations Then   'Ticket #16189
            addNode xparent, lStr("Associations"), "key231", "reports"
        End If
        If gSec_Rpt_Master_Education_Seminars Then  'Ticket #16189
            addNode xparent, "Continuing Education", "key232", "reports"
        End If
        If gSec_Rpt_Master_Formal_Education Then    'Ticket #16189
            addNode xparent, "Formal Education", "key233", "reports"
        End If
        If gSec_Rpt_Languages Then  'Ticket #16189
            addNode xparent, "Languages", "key234", "reports"
        End If
        '7.9 - Enhancement - For all the clients
        'Friesens Corporation - Ticket #16189
        'If glbCompSerial = "S/N - 2279W" Then
            If gSec_Rpt_Req_Course_Hist Then    'Ticket #16189
                addNode xparent, "Required Courses", "key240", "reports"
            End If
        'End If
        If gSec_Rpt_Skills Then 'Ticket #16189
            addNode xparent, "Skills", "key235", "reports"
        End If
        If gSec_Rpt_Succession Then
            addNode xparent, "Succession", "key237", "reports"
        End If
        If gSec_Rpt_Master_Education_Seminars Then  'Ticket #16189
            addNode xparent, "Training Matrix", "key236", "reports"
        End If
        If gSec_Rpt_Training_Plan Then  'Ticket #21709
            addNode xparent, "Training Plan", "key23a", "reports"
        End If
        If gSec_Rpt_GapAnalysis Then 'Ticket #16189
            addNode xparent, "Gap Analysis", "key238", "reports"
        End If
        If gSec_Rpt_User_Defined_Table Then 'Ticket #16189
            addNode xparent, lStr("User Defined Table"), "key239", "reports"
        End If
    End If

xparent = "key80"   'Friesens Reports
    If Not glbtermopen Then
        'Friesens Corporation - Ticket #16189
        If glbCompSerial = "S/N - 2279W" Then
            If gSec_Rpt_Friesens_IWantToKnowYou Then   'Ticket #16189
                addNode xparent, "I Want You to Know...", "key801", "reports"
            End If
            If gSec_Rpt_Friesens_ITHireForm Then   'Ticket #16189
                addNode xparent, "IT Hire", "key802", "reports"
            End If
            If gSec_Rpt_Friesens_ITNoticeOfChange Then   'Ticket #16189
                addNode xparent, "IT Notice of Change", "key803", "reports"
            End If
            If gSec_Rpt_Friesens_NoticeOfChange Then   'Ticket #16189
                addNode xparent, "Notice of Change", "key804", "reports"
            End If
            If gSec_Rpt_Friesens_PerfImproveActionPlan Then   'Ticket #16189
                addNode xparent, "Performance Improvement Action Plan", "key805", "reports"
            End If
            If gSec_Rpt_Friesens_PerformanceReviewRpt Then   'Ticket #16189
                addNode xparent, "Performance Review", "key806", "reports"
            End If
            If gSec_Rpt_Friesens_SeparationRpt Then   'Ticket #16189
                addNode xparent, "Separation HR Acct", "key807", "reports"
            End If
            If gSec_Rpt_Friesens_TerminationRpt Then   'Ticket #16189
                addNode xparent, "Termination HR Acct", "key808", "reports"
            End If
            If gSec_Rpt_Friesens_UpdateMeetingRpt Then   'Ticket #16189
                addNode xparent, "Update Meeting", "key809", "reports"
            End If
            If gSec_Rpt_Friesens_WarningRpt Then   'Ticket #16189
                addNode xparent, "Warning Report", "key810", "reports"
            End If
        End If
    End If
xparent = "key25"
    If Not glbtermopen Then
        If glbWFC Then
            addNode xparent, "Accident Cost Analysis", "key2512", "reports"
        End If
        If gSec_Rpt_Heatlh_Safety Then  'Ticket #16189
            addNode xparent, "Body Site", "key251", "reports"
            addNode xparent, "Day of Week", "key252", "reports"
            addNode xparent, "Employee", "key253", "reports"
            addNode xparent, "Employee Trends", "key254", "reports"
            addNode xparent, "Employee/WSIB Cost Report", "key2510", "reports"
            addNode xparent, "Company Associated Cost Report", "key2513", "reports"
            addNode xparent, "Experience", "key255", "reports"
            addNode xparent, "Incident Type", "key256", "reports"
            addNode xparent, "Injury Code", "key257", "reports"
            addNode xparent, "Plant Area", "key258", "reports"
            addNode xparent, "Shift", "key259", "reports"
        End If
        'addNode xparent, "Total Cost Summary", "key2511", "reports" 'REmoved by Bryan on 2/Dec/05
    End If
    
xparent = "key29"
    If Not glbtermopen Then
        addNode xparent, "Custom Reports", "key293", "reports"
        If glbCompSerial = "S/N - 2257W" Then   'Hamilton C.C.A.S
            addNode xparent, "Hamilton CCAS Custom Reports", "key293a", "reports"
        End If
    End If
    If Not glbtermopen Then
        If gSec_Rpt_Master_Passwords Then   'Ticket #16189
            addNode xparent, "Security Master", "key291", "reports"
        End If
    End If
    If gSec_Rpt_Master_Table_Codes Then 'Ticket #16189
        addNode xparent, "Table Master", "key292", "reports"
    End If
      
    If glbWFC Then
        If gSec_Rpt_Master_Job Then
            xparent = "key27wfc"
            addNode xparent, "Positions/Skills/Evaluation", "key27", "reports"
            If gSec_Inq_BudgetedPos Then
                addNode xparent, "Budgeted Position Report", "key311m", "reports"
                addNode xparent, "Position Control Table Exports", "key312m", "reports"
            End If
        End If
    End If
    
xparent = "key31"
    If Not glbtermopen Then
        If gSec_Rpt_Master_Benefits Then    'Ticket #16189
            addNode xparent, "Benefit Group Change", "key311a", "reports"
            addNode xparent, "Benefits/Beneficiaries", "key311", "reports"
        End If
        If glbCompSerial = "S/N - 2369W" Then   'Ticket #16189
            addNode xparent, "Bonus Report", "tsCBSheet", "reports" 'added by Bryan 30/Nov/05 Ticket#9721
        End If
        If gSec_Rpt_Cost_Of_Employment Then 'Ticket #16189
            addNode xparent, "Cost of Employment", "key312", "reports"
        End If
        If gSec_Rpt_Master_DolEnt Then  'Ticket #16189
            addNode xparent, "Dollar Entitlements", "key313", "reports"
        End If
        If glbGP Or glbCompSerial = "S/N - 2259W" Then 'added by George Feb 23,2006 Ticket#9965 Great Plains/County of Oxford
            'Ticket #22682 - Jerry asked to change to Salary Master report security
            If gSec_Rpt_Master_Salaries Then
            'If gSec_Inq_Salary Then 'Ticket #16189
                addNode xparent, "GP Salary Posting Report", "gpPosting", "reports"
            End If
        End If
        If gSec_Rpt_Master_OtherEarn Then   'Ticket #16189
            addNode xparent, "Other Earnings", "key314", "reports"
        End If
        If gSec_Rpt_Payroll_Trans Then   'Ticket #16189
            addNode xparent, "Payroll Transactions", "keyPayTranRpt", "reports"
        End If
        If gSec_Rpt_Profit_Sharing Then   'Ticket #20052 Franks 07/26/2011
            addNode xparent, "Profit Sharing", "keyProfitSharingRpt", "reports"
        End If
        If gSec_Rpt_Red_Circled Then   'Ticket #20648 Franks 09/26/2011
            addNode xparent, "Red Circled Report", "keyRedCircledRpt", "reports"
        End If
        If gSec_Rpt_Master_Salaries Then    'Ticket #16189
            addNode xparent, "Salary Master", "key315", "reports"
        End If
        If gSec_Rpt_Salary_Performance Then 'Ticket #16189
            If glbCompSerial = "S/N - 2368W" Then
                addNode xparent, lStr("Performance") & " Review", "key316", "reports"
            Else
                addNode xparent, "Salary " & lStr("Performance") & " Review", "key316", "reports"
            End If
        End If
        'Casey House - Ticket #15276
        If glbCompSerial = "S/N - 2214W" Then
            If gSec_Rpt_Entitlements Then   'Ticket #16189
                addNode xparent, "Salary/Vacation Increase Report", "key70", "reports"
            End If
        End If
        
        If glbCompSerial = "S/N - 2279W" Then ' Friesen tkt# 10844
            'If gSec_Rpt_Salary_Performance Then 'Ticket #16189
            If gSec_Rpt_Staff_Profile Then 'Ticket #27795 - Friesens Corporation
                addNode xparent, "Staff Profile", "key317", "reports"
            End If
            If gSec_Rpt_Temp_CrossTraining Then 'Ticket #16189
                addNode xparent, "Temporary/Cross Training Assignment", "key318", "reports"
            End If
        End If
                
    End If
    
'Hemu - 06/02/2004 - Begin
xparent = "key32a"
    If Not glbtermopen Then
    'Hemu - Begin 07/23/2004 - For Surrey Place Only - Indicator reports - after they have
    'successfully tested these reports will be available for others too.
    'If glbCompSerial = "S/N - 2347W" Then
        If glbCompSerial = "S/N - 2369W" Then
            If gSec_Rpt_Manpower_Plan Then  'Ticket #16189
                addNode xparent, "Daily Manpower Update", "key32a10", "reports"  ' added by Bryan 9/Sep/05 Ticket #9235
            End If
        End If
        
        If gSec_Rpt_External_Hire Then  'Ticket #16189
            addNode xparent, "External Hire Rate", "key32a5", "reports"
        End If
        
        If glbCompSerial = "S/N - 2369W" Then
            addNode xparent, "Health & Safety Sheet", "tsHSSheet", "reports" 'added by Bryan 09/Nov/05 Ticket#9720
        End If
        If gSec_Rpt_Internal_Hire Then  'Ticket #16189
            addNode xparent, "Internal Transfers to Total Hires Ratio", "key32a6", "reports"
        End If
        If gSec_Rpt_Key_Workforce Then  'Ticket #16189
            addNode xparent, "Key Workforce Demographic", "key32a1", "reports"
        End If
        If gSec_Rpt_Manpower_Plan Then  'Ticket #16189
            addNode xparent, "Manpower Plan", "key32a9", "reports" 'added by Bryan 14/07/05 Ticket #8921
        End If
        If glbSQL Or glbOracle Then 'Bryan Ticket#11134 Access don't get this report
            If gSec_Rpt_Paid_Sick Then  'Ticket #16189
                addNode xparent, "Paid Sick Hours Per Eligible Employee", "key32a8", "reports"
            End If
        End If
        If glbCompSerial = "S/N - 2369W" Then 'added by Bryan 08/Nov/05 Ticket#9720
            addNode xparent, "Quarterly Report", "tsQuarter", "reports"
        End If
        If gSec_Rpt_Staff_Management Then   'Ticket #16189
            addNode xparent, "Staff/Management Ratios", "key32a2", "reports"
        End If
        'Hemu - This report will remain for Surrey Place only
        If glbCompSerial = "S/N - 2347W" Or glbCompSerial = "S/N - 2394W" Then 'SPC or St. Johns 'Ticket #15558
            If gSec_Rpt_Master_Job Then 'Ticket #16189
                addNode xparent, "Turnover Rates", "key32a7", "reports"
            End If
        End If
        If gSec_Rpt_WC_Time Then    'Ticket #16189
            addNode xparent, "Workers Compensation (WC) Lost Time Incident Rate", "key32a3", "reports"
        End If
        If gSec_Rpt_WC_Work Then    'Ticket #16189
            addNode xparent, "Workers Compensation (WC) Lost Work Hours Rate", "key32a4", "reports"
        End If
    'End If
        'Hemu - End
    End If
'Hemu - 06/02/2004 - End

xparent = "key3Forms"
    'Ticket #24663 - for all
    If Not glbtermopen Then 'Ticket #25954 Franks 09/04/2014
        If gSec_RptF_Attendance_SignIn Then    'Ticket #16189
            addNode xparent, "Attendance Sign In", "key3F1", "reports"
        End If
    End If
    'Ticket #24663 - Showa only
    If glbCompSerial = "S/N - 2454W" Then
        If gSec_RptF_ATT_Discipline Then    'Ticket #16189
            addNode xparent, "ATT Discipline", "key3F2", "reports"
        End If
        If gSec_RptF_COC_Discipline Then    'Ticket #16189
            addNode xparent, "COC Discipline", "key3F3", "reports"
        End If
    End If

xparent = "key4"
    If Not glbtermopen Then
        If GetMassUpdateSecurities("Attendance_His_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Attendance History", "key32", "mass"
        End If
        If GetMassUpdateSecurities("Attendance_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Attendance Master", "key33", "mass"
        End If
        If GetMassUpdateSecurities("Benefits_MassUpdate", glbUserID) Then   'Ticket #16189
            addNode xparent, "Benefits Master", "key34", "mass"
        End If
        If gSec_Mass_Codes Then 'Ticket #16189
            addNode xparent, "Codes", "key35", "mass"
        End If
        If GetMassUpdateSecurities("Education_Seminars_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Continuing Education", "key36", "mass"
        End If
        
        'Release 8.1 - Ticket #27244: Import document Attachment, Assign Document Type Info. under Mass Updates menu
        If GetMassUpdateSecurities("ImpAttachment_MassUpdate", glbUserID) Then
            addNode xparent, "Document Attachment", "key199", "mass"
            'addNode xparent, "Import Attachment", "key199", "mass"
        End If
        
        If GetMassUpdateSecurities("Other_Entitlements_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Dollar Entitlements", "key37", "mass"
        End If
        If glbLinamar Then
            If gSec_Inq_DoorAccess Then 'Ticket #16189
                addNode xparent, "Door Access", "dooraccess", "mass"
            End If
        End If
        If GetMassUpdateSecurities("Emergency_Leave_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Emergency Leave", "EMLSETUP", "mass"
        End If
        'Ticket #21562 - Begin - Making a menu item and having Active Employee and Terminated Employee options
        'to change the Employee #
        If GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Employee Number", "key38ab", "mass"
        End If
        'Ticket #21562 - opening for all
        'If glbCompSerial = "S/N - 2407W" Then ' Farmer's Mutual Mostafa
            'If GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbUserID) Then 'Ticket #16189
            '    addNode xparent, "Terminated Employee Number", "key38a", "mass"
            'End If
        'End If
        'Ticket #21562 - End
        
        'Ticket #28457 - City of Niagara Falls - Employee/Position Mass Update
        If glbCompSerial = "S/N - 2276W" Then
            If GetMassUpdateSecurities("Job_Master_MassUpdate", glbUserID) And GetMassUpdateSecurities("Salary_MassUpdate", glbUserID) Then
                addNode xparent, "Employee Position", "key42a", "mass"
            End If
        End If
        
        If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
            addNode xparent, "Entitlements Standard", "key39", "mass"
            addNode xparent, "VitalAire Entitlements", "key39a", "mass"
        Else
            addNode xparent, "Entitlements", "key39", "mass"
        End If
        If gSec_Inq_Terminations Then 'Ticket #16303
            If glbWFC Then
                addNode xparent, "Enter Leave", "EnterLeave", "mass"
            End If
        End If
        
        If GetMassUpdateSecurities("Follow_Ups_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, lStr("Follow-ups"), "key40", "mass"
        End If
                        
        'Release 8.0 - Ticket #24361: Add Email Address import under Mass Updates menu
        If GetMassUpdateSecurities("EmailLoad_MassUpdate", glbUserID) Then
            addNode xparent, "Import Email Address", "key200", "mass"
        End If
        
        If glbSQL Or glbOracle Then
            If GetMassUpdateSecurities("Import_Photo_MassUpdate", glbUserID) Then 'Ticket #16189
                '8.0 - Ticket #22682 - Export to Folder or Delete Photos from HR_PHOTO
                'addNode xparent, "Import Photo", "key40-1", "mass"
                addNode xparent, "Maintain Photos", "key40-1", "mass"
            End If
        End If
        
        If GetMassUpdateSecurities("Other_Earnings_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Other Earnings", "key41", "mass"
        End If
        If GetMassUpdateSecurities("OvertimeMaster_MassUpdate", glbUserID) Then 'Ticket #16189
            If glbCompSerial = "S/N - 2425W" Then   'Ticket #18223 - Four Villages CHC
                addNode xparent, "Extra Time Master", "key61", "mass"
            Else
                addNode xparent, "Overtime Master", "key61", "mass"
            End If
        End If
        If GetMassUpdateSecurities("Job_Master_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Position Master", "key42", "mass"
        End If
        If gSec_Upd_Position And gSec_Inq_Performance Then  'Ticket #16189
            addNode xparent, "Reporting Authority", "key43", "mass"
        End If
        If GetMassUpdateSecurities("Salary_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Salary Master", "key44", "mass"
        End If
        
        If glbCompSerial = "S/N - 2214W" Then   'Casey House - Ticket #15276
            If gSec_Inq_Salary Then 'Ticket #16189
                addNode xparent, "Salary Increase Rule", "key62", "mass"
            End If
        End If
        
        If glbCompSerial = "S/N - 2420W" Then   'Macaulay - Ticket #25015
            If gSec_Upd_Basic Then   'Ticket #16189
                addNode xparent, "Seniority Date Calculation", "key110", "mass"
            End If
        End If
        
        If gSec_Upd_Banking Then    'Ticket #16189
            addNode xparent, "TD1 Dollar/Code", "key45", "mass"
        End If
        If glbCompSerial <> "S/N - 2259W" Then  'County of Oxford
            If gSec_Upd_Terminations Then   'Ticket #16189
                addNode xparent, "Terminations", "key46", "mass"
            End If
        End If
        
        If glbCompSerial = "S/N - 2214W" Then   'Casey House - Ticket #15276
            If gSec_Inq_Salary Then 'Ticket #16189
                addNode xparent, "Vacation Increase Rule", "key63", "mass"
            End If
        End If
        
        If GetMassUpdateSecurities("Work_Schedule_MassUpdate", glbUserID) Then
            addNode xparent, "Work Schedule", "key98", "mass"
        End If
        
        If glbBurlTech Then
            addNode xparent, "BTI Attendance Entitlement", "key47d", "mass"
            xparent = "key47d"
            'addNode xparent, "Quarter End", "key47d1", "mass" 'Ticket #12378
            If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
                addNode xparent, "Year End Carryover", "key47d2", "mass"
            End If
            'addNode xparent, "Year End Reduction For BD", "key47d3", "mass"
            'addNode xparent, "Year End Reduction For Non BD", "key47d4", "mass"
        End If
    End If
    
'Ticket #21562 - Begin
'Making a menu item and having Active Employee and Terminated Employee options
'to change the Employee #
If Not glbtermopen Then
    xparent = "key38ab"
        If GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbUserID) Then 'Ticket #16189
            addNode xparent, "Active", "key38", "mass"
        End If
        'Ticket #21562 - opening for all but not available for Linamar, Essex County Library and VitalAire,
        'as they have custom logic for Active Employee # change and nothing defined for Term Employee # change.
        'They have not asked for it but when they do we will have to look into their logic before opening up.
        If Not glbLinamar And glbCompSerial <> "S/N - 2296W" And glbCompSerial <> "S/N - 2380W" Then
        'If glbCompSerial = "S/N - 2407W" Then ' Farmer's Mutual Mostafa
            If GetMassUpdateSecurities("EmployeeNo_MassUpdate", glbUserID) Then 'Ticket #16189
                addNode xparent, "Terminated", "key38a", "mass"
            End If
        'End If
        End If
End If
'Ticket #21562 - End

'Release 8.1 - Ticket #27244: Document Attachment related Menu: Import Attachment, Update Document Type Info.
If Not glbtermopen Then
    xparent = "key199"
        If GetMassUpdateSecurities("ImpAttachment_MassUpdate", glbUserID) Then 'Ticket #16189
            'Release 8.1 - Ticket #27244: Update Document Type Info.
            addNode xparent, "Attach Document Type", "key199b", "mass"
            
            'Release 8.1 - Ticket #27244: Import document Attachment
            addNode xparent, "Import Attachment Files", "key199a", "mass"
        End If
End If

xparent = "key39"
    If Not glbtermopen Then
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            addNode xparent, "Clear Accrual File", "key397", "mass"
        End If
        If gSec_Inq_Holiday Then    'Ticket #16189
            addNode xparent, "Holiday Master", "key391", "mass"
        End If
        If GetMassUpdateSecurities("Hrly_Entitlements_MassUpdate", glbUserID) Then  'Ticket #16189
            addNode xparent, "Hourly Entitlement Master", "key394", "mass"
            'addNode xparent, "Hours Based Entitlement Master", "key340", "mass"
        End If
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            addNode xparent, "Pension Entitlement Master", "keyPPct", "mass"
            'addNode xparent, "Rollover Entitlements", "key396", "mass"
            addNode xparent, "Rollover Entitlements", "key396Rollover", "mass"
            
            'Ticket #17924
            'addNode xparent, "Rollover Hourly Entitlements", "key396a", "mass"
        End If
        If glbCompSerial <> "S/N - 2380W" Then 'VitalAire Ticket #13979
            If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
                If glbCompSerial <> "S/N - 2418W" Then 'Ticket #17786
                    addNode xparent, "Sick Entitlement Master", "key393", "mass"
                End If
                
                'Ticket #29230 - Daily Vacation Entitlement
                '              - Only display this screen if Vacation Enttitlement Earned is Daily and Vacation Entitlement Outstanding Based On = 1
                If glbCompEntVacDaily And glbCompEntVac$ = "D" And glbEntOutStanding$ = "1" Then
                    addNode xparent, "Daily Vacation Accrual Master", "key392b", "mass"
                    'addNode xparent, "Daily Accrual Skipped Log", "key392c", "mass"
                Else
                    'Ticket #29230 - Daily Vacation Entitlement
                    'If Daily Vacation Accrual option selected on the Company Master then they do not get the Vacation Entitlement Master option. It's one or the other.
                    addNode xparent, "Vacation Entitlement Master", "key392", "mass"
                End If
                
                'Ticket #25943 - Hours Based Entitlement & Vacation Pay % for Hours Based Vacation Entitlement
                addNode xparent, "Hours Based Vacation Entitlement Master", "key399", "mass"
                addNode xparent, "Vacation Pay Percentage Master", "key398", "mass"

                If glbCompSerial = "S/N - 2460W" Then   'Ticket #26154 - Oshawa Public Libraries
                    addNode xparent, "Vacation By Pay Period", "key392a", "mass"
                End If
            End If
        End If
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            'addNode xparent, "Zero Out Entitlements", "key395", "mass"     'Ticket #17924 - Submenu added instead
            addNode xparent, "Zero Out Entitlements", "key395ZeroOut", "mass"
            
            'Ticket #17924
            'addNode xparent, "Zero Out Hourly Entitlements", "key395a", "mass"
        End If
    End If

xparent = "key396Rollover"
    If Not glbtermopen Then
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            If glbCompSerial = "S/N - 2425W" Then 'Four Villages (Ticket #19998)
                addNode xparent, "Vacation, Sick and Extra Time", "key396", "mass"
            Else
                addNode xparent, "Vacation, Sick and Overtime", "key396", "mass"
            End If
            If glbSQL Then
                addNode xparent, "Hourly Entitlements", "key396a", "mass"
            End If
        End If
    End If
    
xparent = "key395ZeroOut"
    If Not glbtermopen Then
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            If glbCompSerial = "S/N - 2425W" Then 'Four Villages (Ticket #19998)
                addNode xparent, "Vacation, Sick and Extra Time", "key395", "mass"
            Else
                addNode xparent, "Vacation, Sick and Overtime", "key395", "mass"
            End If
            If glbSQL Then
                addNode xparent, "Hourly Entitlements", "key395a", "mass"
            End If
        End If
    End If
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #13979
xparent = "key39a"
    If Not glbtermopen Then
        If GetMassUpdateSecurities("Entitlements_MassUpdate", glbUserID) Then   'Ticket #16189
            addNode xparent, "Annual Entitlement Master", "key393b", "mass"
            addNode xparent, "Pay Period - Current Accrued Master", "key392", "mass"
            addNode xparent, "Sick Entitlement Master", "key393", "mass"
            addNode xparent, "Year End - Current Accrued", "key393a", "mass"
        End If
    End If
End If

If glbWFC And glbWFC_IncentivePlanFlag Then 'Ticket #29013 Franks 08/23/2016
    If Not glbtermopen Then
        xparent = "keyIP"
        If gSec_WFC_IPExchangeRate Then ' gSec_Upd_Company Then
            addNode xparent, "Currency Exchange Rate", "keyIP1", "applicants"
        End If
        If gSec_WFC_IPIncentiveFactors Then ' gSec_Upd_Company Then
            addNode xparent, "Company Incentive Factors", "keyIP2", "applicants"
        End If
        If gSec_WFC_IPCreateSpreadsheet Then ' gSec_Upd_Company Then
            addNode xparent, "Create Incentive Plan Spreadsheet", "keyIP3", "applicants"
        End If
        If gSec_WFC_IPImportSpreadsheet Then 'gSec_Upd_Company Then
            addNode xparent, "Import Incentive Plan Spreadsheet", "keyIP4", "applicants"
        End If
        If gSec_WFC_IPUpdateEarnings Then 'gSec_Upd_Company Then
            addNode xparent, "Update info:HR Other Earnings", "keyIP5", "applicants"
        End If
        If gSec_WFC_IPPreparePayrollFile Then 'gSec_Upd_Company Then
            addNode xparent, "Prepare Payroll Transaction File", "keyIP6", "applicants"
        End If
        If gSec_WFC_IPPrintSpreadsheet Then 'gSec_Upd_Company Then
            addNode xparent, "Print Incentive Plan Spreadsheet", "keyIP7", "applicants"
        End If
        If gSec_WFC_IPPrintLetter Then  'gSec_Rpt_Master_Salaries Then
            addNode xparent, "Print Employee Incentive Letter", "keyIP8", "applicants"
        End If
        'Ticket #29633 Franks 01/06/2017 - Remove from Incentive menu.
        'If gSec_Upd_Company Then
        '    addNode xparent, "Print Scorecard Data", "keyIP9", "applicants"
        'End If
    End If
End If
    
xparent = "key5"
    If Not glbtermopen Then
        If gSec_Inq_Project_Code Then   'Ticket #16189
            addNode xparent, lStr("Account Code Master"), "ProjectCode", "setup"
        End If

        'Ticket #30508 - Applicant Tracking Enhancement
        If mmnu_AppT.Visible Then
            addNode xparent, "Applicant Tracking Setup", "keyAppTSetup", "setup"
            
            xparent = "keyAppTSetup"    'child node
            If gSec_Inq_AppFormDefaults Then    'Ticket #16189
                addNode xparent, "Application Form Defaults", "keyAppT003", "setup"
            End If
            If gSec_Inq_AppFormWorkFlow Then    'Ticket #16189
                addNode xparent, "Application Form Workflow", "keyAppT002", "setup"
            End If
            If gSec_Inq_LettersPosType Then   'Ticket #16189
                addNode xparent, "Letter by Position Type", "keyAppT001", "setup"
            End If
            
            xparent = "key5"    'reset back to the original parent node
        End If

        If gSec_Inq_AttendCode_Matrix Then
             addNode xparent, "Attendance Code Matrix", "key58b", "setup"
        End If

        If glbCompSerial = "S/N - 2233" Then
            If gSec_Inq_Attendance_Group_Code_Matrix Then
                 addNode xparent, "Attendance Group Matrix", "key58a", "setup"
            End If
        End If

        If gSec_Inq_Audit Then  'Ticket #16189
            'addNode xparent, "Audit Logs", "auditlog", "setup"
            
            If glbCompSerial = "S/N - 2241W" Then ' Granite Club Ticket #16017
                addNode xparent, "Attendance Audit", "key47e", "setup"
            Else
                'Release 8.0 - Ticket #22682
                'addNode xparent, "Attendance Audit Master", "key47e", "setup"
            End If
            addNode xparent, "Audit Master", "key47", "setup"
        End If
        
        addNode xparent, "Benefit Group Setup", "key47BenGroup", "setup"
        
        If glbWFC Then
            If gSec_Inq_Departments Then    'Ticket #16189
                addNode xparent, ("Bonus Reporting No Master"), "key47ad", "setup"
            End If
        End If
        If gSec_Inq_BudgetedMP Then     'Ticket #16189
            addNode xparent, "Budgeted Manpower", "key60", "setup" 'added by Bryan 12/07/05 Ticket #8922
        End If
        If glbCompSerial = "S/N - 2351W" Then ' For Burlington Tech.
            If gSec_Inq_Charge_Code Then    'Ticket #16189
                addNode xparent, "Charge Code Master", "ChargeCode", "setup"
            End If
        End If
        If gSec_Inq_Company Then    'Ticket #16189
            addNode xparent, "Company Master", "key48", "setup"
        End If
        If gSec_CompanyPreference Then  'Ticket #16189
            addNode xparent, "Company Preference", "key48C", "setup"
        End If
        If gSec_Inq_CourseCodeMaster Then    'Ticket #16189
            addNode xparent, "Course Code Master", "key48D", "setup"
        End If
        If glbCompSerial = "S/N - 2351W" Then ' For Burlington Tech.
            If gSec_Inq_Company Then    'Ticket #16189
                addNode xparent, "Counseling Target Setup", "key48B", "setup"
            End If
        End If
        If gSec_Inq_CustomReport Then   'Ticket #16189
            addNode xparent, "Custom Reports Master", "key49", "setup"
        End If
        
        'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
        If glbCompSerial = "S/N - 2382W" Then
            If gSec_Inq_CounselAudit Then
                addNode xparent, lStr("Counseling Audit Master"), "key100", "setup"
            End If
        End If
                
        If gSec_Inq_DashboardRule Then 'Ticket #22541
            addNode xparent, "Dashboard Setup", "DashboardSetup", "setup"
        End If
        
        If gSec_Inq_Departments Then    'Ticket #16189
            addNode xparent, lStr("Department Master"), "key50", "setup"
        End If
        
        'Ticket #25746 - Department/GL Number Matrix
        If gSec_Inq_DeptGL_Matrix Then
             addNode xparent, lStr("Department") & "/" & lStr("G/L") & " Matrix", "key58e", "setup"
        End If
        
        If gSec_Inq_Divisions Then  'Ticket #16189
            addNode xparent, lStr("Division Master"), "key51", "setup"
        End If
        If glbWFC Then
            If glbPlantCode = "WHBY" Then
                If Not gSec_Inq_Master_Table_Exists("CERE") Then   'Ticket #16189
                    addNode xparent, "Disciplinary Steps", "key510", "setup"
                End If
            End If
        End If
        
        'If glbUS Then
        If glbUS Or glbCountry = "ALL" Then 'Ticket #18664
            'addNode xparent, "Affirmative Action", "key47a", "setup"
            addNode xparent, "EEO", "key47a", "setup"
        End If
        
        addNode xparent, "Employment Equity", "key51a", "setup"

        'If gSec_Inq_Label Then  'Ticket #16189
        '    addNode xparent, "Employee Flags", "EmpFlags", "setup"
        'End If
        
        If glbVadim Then
            If gSec_Inq_Machine Then    'Ticket #16189
                addNode xparent, "Equipment Master", "Machine", "setup"
            End If
        End If
        
        'Ticket #22682 - Release 8.0 - Follow Up Code Email Matrix
        If gSec_Inq_FollowUpEmail_Matrix Then
             addNode xparent, lStr("Follow-ups Code Email Matrix"), "key58d", "setup"
        End If
                
        'For WSIB Form 7
        If glbWSIBModule Then
            addNode xparent, "Form 7", "Form7", "setup"
            
'            If gSec_Inq_HSW7CmpMst Then     'Ticket #16189
'                addNode xparent, "Form 7 Employer Information", "Key121a", "setup"
'            End If
'
'            addNode xparent, "Form 7 Employee Type Matrix", "key65", "setup"
        End If
        
        If gSec_Inq_Ledgers Then    'Ticket #16189
            addNode xparent, lStr("G/L Master"), "key52", "setup"
        End If
        
        addNode xparent, "Help Descriptions", "key5help", "setup"
        If gSec_HelpDescSetup Then  'Ticket #16189
            addNode "key5help", "Corrective Action Code", "key5help1", "setup"
        End If
        
        If glbCompSerial = "S/N - 2380W" Then ' For VitalAire Canada Inc. Ticket #26233 Franks 11/20/2014
            If gSec_Upd_DoorAccess Then
                addNode xparent, "Job Classification Tables", "key5JobClass", "setup"
                addNode "key5JobClass", "Job Family", "key5JobClass1", "setup"
                addNode "key5JobClass", "Sub-Job Family", "key5JobClass2", "setup"
                addNode "key5JobClass", "Group Jobs", "key5JobClass3", "setup"
            End If
        End If
        
        'Frank 06/18/2004 ticket# 6382
        'If mmnu_Opus_Payroll.Enabled = True Then
        If xIntellisolMatrix Then
            addNode xparent, "Intellisol Matrix", "key53", "setup"
        End If
        If gSec_Inq_Label Then  'Ticket #16189
            addNode xparent, "Label Master", "key54", "setup"
        End If
        If glbLinamar Then
            addNode xparent, "Linamar Codes Master", "keyLinamarCode", "setup"
        End If
        'If glbLambton Then
        '    'If App.Path = "C:\SSWORK\IHR72" Or App.Path = "U:\HR Systems VB6\IHR 7.2" Then
        '        addNode xparent, "Lambton Matrix", "key550", "setup"
        '        addNode xparent, "Lambton Total Hours Worked", "key551", "setup"
        '    'End If
        'End If
                
        If Not glbVadim Then
            If gSec_Inq_Machine Then    'Ticket #16189
                addNode xparent, lStr("Machine # Master"), "Machine", "setup"
            End If
        End If
        
        If glbSQL Or glbOracle Then
            If gSec_MultiDataSourceSetup Then   'Ticket #16189
                addNode xparent, "Multiple Data Sources", "MultipleDS", "setup"
            End If
        End If
        If glbWFC Then
            If (glbPlantCode = "MISS" Or glbPlantCode = "TROY") Then
                'addNode xparent, "Market Line Master", "key550", "setup"
            End If
        End If
        If gSec_Inq_New_Hire Then   'Ticket #16189
            addNode xparent, "New Hire Procedure", "key55", "setup"
        End If
        
        'Ticket #24655 - Wellington-Dufferin-Guelph Public Health - On Call Hours
        If glbCompSerial = "S/N - 2411W" Then
            If gSec_Inq_OnCallHours Then
                addNode xparent, "On Call Hours", "OnCallHrs", "setup"
            End If
        End If
        
        'Ticket #25922 - OHRS Reporting for CHC
        If gSec_Inq_OHRSDepartments Then
            addNode xparent, "OHRS Department Master", "key50a", "setup"
        End If
        
        If Not glbLinamar Then
            If gSec_Inq_PayPeriod_Master Then   'Ticket #16189
                addNode xparent, "Pay Period", "key56-A", "setup"   'Ticket #29617 - Made it parent
                'addNode xparent, "Pay Period Master", "key56-0", "setup"
            End If
        End If
        
        If gSec_Matrix Then 'Ticket #16189
            addNode xparent, "Payroll Matrix", "key56", "setup"
        End If
        
        'Friesens Corporation - Ticket #16189
        If glbCompSerial = "S/N - 2279W" Then
            If gSec_Inq_Performance And gSec_Inq_Job_Master Then 'Ticket #16189
                addNode xparent, "Position Group/Performance Category Link", "key64", "setup"
            End If
        End If
        
        If gSec_Province Then
            addNode xparent, "Province/State Master", "key57", "setup"
        End If
        
        addNode xparent, "Root Cause Links", "key517", "setup"
        
        'Ticket #18335 Frank 04/12/2010, Jerry asked to use OHS Cause Code security
        If gSec_Inq_Master_Table_Exists("ECCA") Then
            'If gSec_Inq_Company Then    'Ticket #16189
            If glbLinamar Then  'Ticket #14620
                addNode "key517", "Substandard Act", "key5171", "setup"
                addNode "key517", "Substandard Condition", "key5172", "setup"
                addNode "key517", "Personal Factor", "key5173", "setup"
            Else
                addNode "key517", "Type of Event", "key5171", "setup"
                addNode "key517", "Immediate / Direct Causes", "key5172", "setup"
                addNode "key517", "Basic / Underlying Causes", "key5173", "setup"
            End If
        End If
        
        If gSec_Inq_SalDist Then    'Ticket #16189
            addNode xparent, lStr("Salary Distribution Master"), "SalDist", "setup"
        End If
        
        addNode xparent, "Security", "key59", "setup"
        
        If gSec_Inq_Master_Table.count <> 0 Then    'Ticket #16189
            addNode xparent, "Table Master", "key58", "setup"
            If gSec_Inq_SAMTableMasterLinks Then 'Ticket #21581 Franks 02/14/2012
            'If glbSamuel Then 'Ticket #21106 Franks 11/04/2011
                addNode xparent, "Table Master Edit Links", "key58c", "setup"
            End If
        End If
        
        If glbWFC Then 'Ticket #15248
            If gSec_Inq_Company Then    'Ticket #16189
                addNode xparent, "Termination Reason Link", "key59b", "setup"
                addNode xparent, "Work Flow Master", "key59c", "setup" 'Ticket #16395
            End If
        End If
        
        If gSec_Inq_WorkSchRule Then     'Ticket #22220
            addNode xparent, "Work Schedule Rules", "WorkSchRule", "setup"
        End If
        
        If glbWHSCC Then
            If gSec_Inq_WHSCC_USB Then  'Ticket #16189
                addNode xparent, "Union Sick Bank", "key59a", "setup"
            End If
        End If
        'If glbUS Then
        If glbUS Or glbCountry = "ALL" Then 'Ticket #18664
            If glbSQL Then 'Ticket #18790 sql only
                xparent = "key47a"
                If gSec_Inq_AffirmAction_Data Then
                    addNode xparent, "Data Maintenance", "key47b", "setup"
                End If
                If gSec_Rpt_AffirmAction Then
                    addNode xparent, "Reports", "key47c", "setup"
                End If
                If gSec_Inq_AffirmAction_Purge Then
                    addNode xparent, "Purge Applicants Records", "key47g", "setup"
                End If
            End If
        End If
    End If

'If gSec_Inq_Audit Then  'Ticket #16189
'    xparent = "auditlog"
'        If glbCompSerial = "S/N - 2241W" Then ' Granite Club Ticket #16017
'            addNode xparent, "Attendance Audit", "key47e", "setup"
'        Else
'            'Release 8.0 - Ticket #22682
'            addNode xparent, "Attendance Audit Master", "key47e", "setup"
'        End If
'        addNode xparent, "Audit Master", "key47", "setup"
'
'        'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
'        If glbCompSerial = "S/N - 2382W" Then
'            If gSec_Inq_CounselAudit Then
'                addNode xparent, lStr("Counseling Audit Master"), "key100", "setup"
'            End If
'        End If
'End If


'Ticket #22682 - Release 8.0 - Break down Company Preference
If gSec_CompanyPreference Then
    xparent = "key48C"
    If Not glbtermopen Then
        addNode xparent, "General", "key48C_General", "setup"
        addNode xparent, "Email Notifications", "key48C_EmailNotification", "setup"
        addNode xparent, "File Locations", "key48C_FileLocation", "setup"
    End If
End If

If gSec_Inq_Label Then  'Ticket #16189
    xparent = "key54"
        If Not glbtermopen Then
            If gSec_Rpt_Employee_Labels Then    'Ticket #16189
                addNode xparent, "Basic Information", "key54_LblBasicInfo", "setup"
                xparent = "key54_LblBasicInfo"
                    addNode xparent, "Demographics", "key54_Lbl1", "setup"
                    addNode xparent, "Status/Dates", "key54_Lbl2", "setup"
                    addNode xparent, "Dependents", "key54_Lbl3", "setup"
                    addNode xparent, "Banking Information", "key54_Lbl4", "setup"
                    addNode xparent, "Other Information", "key54_Lbl5", "setup"
                    addNode xparent, "Employee Flags", "key54_Lbl6", "setup"
                    
                    'Ticket #25015 - Macaulay: New Additional Payroll ID Data
                    If gSec_Upd_AddPayrollIDData Then
                        addNode xparent, lStr("Additional Payroll ID Data"), "key54_Lbl20", "setup"
                    End If
                xparent = "key54"
                addNode xparent, "Work History/Compensation", "key54_LblWorkHist", "setup"
                xparent = "key54_LblWorkHist"
                    addNode xparent, "Position", "key54_Lbl7", "setup"
                    addNode xparent, "Salary", "key54_Lbl8", "setup"
                    addNode xparent, lStr("Performance"), "key54_Lbl9", "setup"
                xparent = "key54"
                addNode xparent, "Attendance/Entitlements", "key54_LblAttdEnt", "setup"
                xparent = "key54_LblAttdEnt"
                    addNode xparent, "Attendance", "key54_Lbl10", "setup"
                   
                xparent = "key54"
                addNode xparent, "Education/Skills", "key54_LblEduSkl", "setup"
                xparent = "key54_LblEduSkl"
                    addNode xparent, lStr("Associations"), "key54_Lbl11", "setup"
                    addNode xparent, "Continuing Education", "key54_Lbl12", "setup"
                    addNode xparent, lStr("User Defined Table"), "key54_Lbl13", "setup"
                
                xparent = "key54"
                    addNode xparent, lStr("Follow-ups"), "key54_Lbl14", "setup"
                    addNode xparent, lStr("Counseling"), "key54_Lbl15", "setup"
                    addNode xparent, lStr("Comments"), "key54_Lbl16", "setup"
                                
                If glbWFC Then 'Ticket #26254 Franks 12/09/2014
                    addNode xparent, "Job/Position", "key54_LblPosition", "setup"
                Else
                    addNode xparent, "Position", "key54_LblPosition", "setup"
                End If
                xparent = "key54_LblPosition"
                    If glbWFC Then 'Ticket #26254 Franks 12/09/2014
                    addNode xparent, "Job Master", "key54_Lbl21", "setup"
                    End If
                    addNode xparent, "Position Master", "key54_Lbl17", "setup"
            
                'Ticket #22825
                xparent = "key54"
                addNode xparent, "Setup", "key54_LblSetup", "setup"
                xparent = "key54_LblSetup"
                    addNode xparent, "Dashboard Setup", "key54_Lbl18", "setup"
                    'Release 8.0 - Ticket #22682: Add to Label Master
                    addNode xparent, "Province/State Master", "key54_Lbl19", "setup"
            End If
        End If
End If

xparent = "keyLinamarCode"
    If Not glbtermopen And glbLinamar Then
        If gSec_Inq_Master_Table_Exists("BNCD") Then   'Ticket #16189
            addNode xparent, "Benefit", "keyLinamarCode-BNCD", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("HMOP") Then   'Ticket #16189
            addNode xparent, "Home Operation Number", "keyLinamarCode-HMOP", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("HMLN") Then   'Ticket #16189
            addNode xparent, "Home Line", "keyLinamarCode-HMLN", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("HMWC") Then   'Ticket #16189
            addNode xparent, "Home Work Center", "keyLinamarCode-HMWC", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("HMSF") Then   'Ticket #16189
            addNode xparent, "Home Shift", "keyLinamarCode-HMSF", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("EDSE") Then   'Ticket #16189
            addNode xparent, "Operation", "keyLinamarCode-EDSE", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("EDRG") Then   'Ticket #16189
            addNode xparent, "Product Line", "keyLinamarCode-EDRG", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("SHFT") Then   'Ticket #28846 Franks 07/14/2016
            addNode xparent, "Shift", "keyLinamarCode-SHFT", "setup"
        End If
        If gSec_Inq_Master_Table_Exists("EDSK") Then   'Ticket #16189
            addNode xparent, "Skill", "keyLinamarCode-EDSK", "setup"
        End If
        If gSec_Inq_Productline_Operation Then  'Ticket #16189
            addNode xparent, "Product Line/Operation", "keyLinamarCode-RGSE", "setup"
        End If
    End If
        
xparent = "key59"
    If Not glbtermopen Then
        If gSec_ChangeYourPassword Then
            addNode xparent, "Change Your Password", "key591", "setup"
        End If
        If gSec_Inq_Security Then   'Ticket #16189
            addNode xparent, "Security Master", "key592", "setup"
        End If
        If glbLinamar Then
            If gSec_Inq_DoorAccess Then 'Ticket #16189
                addNode xparent, "Door Access", "key592-L1", "setup"
            End If
            If gSec_DoorName Then   'Ticket #16189
                addNode xparent, "Door Name Master", "key592-L2", "setup"
            End If
        End If
        If gSec_Inq_Email_Setup Then 'Ticket #16189    'gSec_Inq_Security Or gSec_Inq_Quick_ESS
            addNode xparent, "Email Setup", "key593", "setup"
        End If
        If gSec_Inq_Email_Setup Then 'Ticket #27127 Franks 05/26/2015
            addNode xparent, "View SMTP Log", "key594", "setup"
        End If
        '7.9 Enhancement - Remove this
        'If gSec_Inq_Quick_ESS Then  'Ticket #16189
        '    addNode xparent, "Quick Setup for ESS", "key595", "setup"
        'End If
    End If
    'If glbCompSerial = "S/N - 2351W" Then ' For Burlington Tech.
    '        xparent = "key48B"
    '        If Not glbtermopen Then
    '            addNode xparent, "Absence Table", "key48B1", "setup"
    '            addNode xparent, "Lateness/Left Early Table", "key48B2", "setup"
    '        End If
    'End If
xparent = "key51a"
    If Not glbtermopen Then
        If gSec_Inq_EmploymentEQT Then  'Ticket #16189
            addNode xparent, "Plan Data", "key51a1", "setup"
            addNode xparent, "Survey Data", "key51a2", "setup"
        End If
        addNode xparent, "Reports", "key51a3", "setup"
    End If
xparent = "key51a3"
    If Not glbtermopen Then
        If gSec_Inq_PayEQT Then 'Ticket #16189
            addNode xparent, "Completed Workforce Surveys", "key51a34", "setup"
            addNode xparent, "Employment Status Analysis", "key51a35", "setup"
            addNode xparent, "Employment Workforce Survey", "key51a36", "setup"
            
            'Ticket #25367 - VitalAire
            If glbCompSerial = "S/N - 2380W" Then
                addNode xparent, "VitalAire Employment Equity", "key51a37", "setup"
            End If
        End If
    End If
xparent = "key47BenGroup"
    If Not glbtermopen Then
        If gSec_BenefitGroupSetup Then  'Ticket #16189
            addNode xparent, "Benefit Costing Details", "key47ac", "setup"
            addNode xparent, "Benefit Group Master", "key47ab", "setup"
            addNode xparent, "Benefit Group Matrix", "key47af", "setup"
            'Ticket #25500 - Goodmans LLP - Benefit Rates
            If glbCompSerial = "S/N - 2290W" Then
                addNode xparent, "Benefit Rates", "key47ai", "setup"
            End If
            
            If glbSQL Then
                addNode xparent, "OMERS Formula", "key47ah", "setup"
            End If
        End If
        'If glbWFC Then  ' 'Ticket #13448
        '    addNode xparent, "Manulife Transaction Rule", "key47ag", "setup"
        'End If

    End If
    
    If glbWFC Then 'Ticket #22285 Franks 07/16/2012
        If gSec_Inq_RetirementProc Then
            xparent = "key156"
            addNode xparent, "ACT->RET Retirement", "key156a", "setup"
            addNode xparent, "ACT->ACP Retired - Working", "key156b", "setup"
            addNode xparent, "ACP->RET Working Retiree Retirement", "key156c", "setup"
        End If
    End If
        
If Not glbtermopen Then
    If Not glbLinamar Then
    xparent = "key56-A"
        If gSec_Inq_PayPeriod_Master Then   'Ticket #16189
            addNode xparent, "Pay Period Master", "key56-0", "setup"
            'Ticket #29617 - Mississaugas of Scugog Island First Nation
            If glbCompSerial = "S/N - 2485W" Then
                addNode xparent, "Close a Pay Period", "key56-1", "setup"
            End If
        End If
    End If
End If

xparent = "keypw6"
    If Not glbtermopen And glbPayWeb Then
        addNode xparent, "Export", "keypw61", "payweb"
        addNode xparent, "Import", "keypw62", "payweb"
        addNode xparent, "Setup", "keypw63", "payweb"
    End If

xparent = "Form7"
    If Not glbtermopen And glbWSIBModule Then
        If gSec_Inq_HSW7CmpMst Then     'Ticket #16189
            addNode xparent, "Employer Information", "Key121a", "setup"
        End If
    
        addNode xparent, "Employee Type Matrix", "key65", "setup"
        addNode xparent, "Filled By", "key66", "setup"
    End If
    
xparent = "keypw61"
    If Not glbtermopen And glbPayWeb Then
        If gSec_Export_Attendance Then  'Ticket #16189
            addNode xparent, "Attendance", "keypw611", "payweb"
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
            addNode xparent, "Initial Data Load", "keypw612", "payweb"
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
            addNode xparent, "Ongoing Employee Transfer", "keypw613", "payweb"
        End If
        'If Not gSec_Upd_Audit Then  'Ticket #16189
        'Ticket #20646 Franks 07/18/2011, fixed the problem which user can't do the reset
        If gSec_Upd_Audit Then   'Ticket #16189
            addNode xparent, "Reset Upload Flags", "keypw614", "payweb"
        End If
    End If
xparent = "keypw62"
    If Not glbtermopen And glbPayWeb Then
        If gSec_Import_Attendance Then  'Ticket #16189
            addNode xparent, "Attendance", "keypw621", "payweb"
        End If
        If gSec_Import_Employee And gSec_Import_Salaries And gSec_Import_Benefits Then  'Ticket #16189
            addNode xparent, "Initial Data Load", "keypw622", "payweb"
        End If
        If gSec_Import_Employee And gSec_Import_Salaries And gSec_Import_Benefits Then  'Ticket #16189
            addNode xparent, "YTD Data", "keypw623", "payweb"
        End If
    End If
xparent = "keypw63"
    If Not glbtermopen And glbPayWeb Then
        addNode xparent, "Payweb FTP ", "keypw631", "payweb"
        If gSec_Matrix Then 'Ticket #16189
            addNode xparent, "Code Matrix", "keypw632", "payweb"
        End If
    End If
xparent = "keyvd7"
    If glbVadim Then
        If gSec_Import_Attendance Then  'Ticket #16189
            addNode xparent, "Attendance Synchronization", "keyvd77", "payweb"
        End If
        If gSec_Import_Attendance Or gSec_Export_Attendance Then    'Ticket #16189
            addNode xparent, "Accrual Class", "keyvd74", "payweb"
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
            addNode xparent, "Code Synchronization", "keyvd75", "payweb"
        End If
        
        'Currently for City of Timmins only
        If glbCompSerial = "S/N - 2375W" Then
            If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
                addNode xparent, "info:HR Code Synchronization", "keyvd75a", "payweb"
            End If
        End If
        
        'Ticket #29122 - New Database Setup and Integration Setup securities
        'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
        If gSec_Inq_IntegrtDBSetup Then
            addNode xparent, "Database Setup", "keyvd72", "payweb"
        End If
        If gSec_Import_Attendance Then  'Ticket #16189
            addNode xparent, "IDL for Accrued Transaction", "keyvd78", "payweb"
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
            addNode xparent, "Vadim Import Table", "keyvd76", "payweb"       'Import Table
            addNode xparent, "info:HR Table Matrix", "keyvd76a", "payweb"    'Code Matrix
        End If
        'Ticket #29122 - New Database Setup and Integration Setup securities
        'If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
        If gSec_Inq_IntegrtSetup Then
            addNode xparent, "Integration Setup", "keyvd71", "payweb"
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
            addNode xparent, "Pay Code Master", "keyvd70", "payweb"
        End If
        If gSec_Inq_Payroll_Category Then   'Ticket #16189
            addNode xparent, "Payroll Category Master", "paycategory", "payweb"
        End If
        If gSec_Export_Benefits Then    'Ticket #16189
            addNode xparent, "Payroll Matrix for Benefit", "paybenmatrix", "payweb"
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Benefits Then  'Ticket #16189
            addNode xparent, "Report", "keyvd73", "payweb"
        End If
        If glbCompSerial = "S/N - 2276W" Then 'City of Niagara Falls
            If gSec_Export_Employee And gSec_Export_Salaries Then   'Ticket #16189
                addNode xparent, "Future Salary Update Report", "keyvd79", "payweb"
            End If
        End If
    End If
xparent = "keyat8"
    If glbAdv Then
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then    'Ticket #16189
            'addNode xparent, "Database Setup", "advDataSetup", "payweb"
            'Ticket #24155 Franks 07/29/2013 - begin
            xTmpFlag = True
            If glbWFC And Not glbWFCFullRights Then xTmpFlag = False
            If xTmpFlag Then addNode xparent, "Database Setup", "advDataSetup", "payweb"
            'Ticket #24155 Franks 07/29/2013 - end
        End If
        If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab - Ticket #15094
            If gSec_Export_Employee And gSec_Export_Attendance And gSec_Import_Attendance Then  'Ticket #16189
                addNode xparent, "Delete Attendance Records", "advDeleteAttnd", "payweb"
            End If
        End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then    'Ticket #16189
            'addNode xparent, "Integration Setup", "advIntgSetup", "payweb"
            'addNode xparent, "Integration Selection", "advSelectionSetup", "payweb"
            'Ticket #24155 Franks 07/29/2013 - begin
            xTmpFlag = True
            If glbWFC And Not glbWFCFullRights Then xTmpFlag = False
            If xTmpFlag Then addNode xparent, "Integration Setup", "advIntgSetup", "payweb"
            If xTmpFlag Then addNode xparent, "Integration Selection", "advSelectionSetup", "payweb"
            'Ticket #24155 Franks 07/29/2013 - end
        End If
        If gSec_Export_Employee And gSec_Export_Attendance And gSec_Import_Attendance Then  'Ticket #16189
            addNode xparent, "Import Attendance", "advImpAttendance", "payweb"
        End If
        If glbWFC Then
            addNode xparent, "Time Bank  Synchronization", "advBankSync", "payweb"
        End If
        If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab - Ticket #15094
            If gSec_Export_Employee And gSec_Export_Attendance And gSec_Import_Attendance Then  'Ticket #16189
                addNode xparent, "Reset Upload Flag", "advResetUpload", "payweb"
            End If
        End If
        'Ticket #20756
        If gSec_Export_Employee And gSec_Export_Table Then
            'addNode xparent, "Initial Data Load", "advExpEmpTblMst_IDL", "payweb"
            'Ticket #24155 Franks 07/29/2013 - begin
            xTmpFlag = True
            If glbWFC And Not glbWFCFullRights Then xTmpFlag = False
            If xTmpFlag Then addNode xparent, "Initial Data Load", "advExpEmpTblMst_IDL", "payweb"
            'Ticket #24155 Franks 07/29/2013 - end
        End If
        'Ticket #21447 Franks 02/15/2012
        If glbCompSerial = "S/N - 2355W" Then 'County of Lambton
            If gSec_Export_Employee Then
                addNode xparent, "Transfer Vacation/Sick Balances", "advExpVacSickBalance", "payweb"
            End If
        End If
    Else
        If glbWFCFullRights Then 'WFC Super users - They have right of Setup of AT
            addNode xparent, "Database Setup", "advDataSetup", "payweb"
            addNode xparent, "Integration Setup", "advIntgSetup", "payweb"
            addNode xparent, "Integration Selection", "advSelectionSetup", "payweb"
            addNode xparent, "Time Bank  Synchronization", "advBankSync", "payweb"
        End If
    End If
    
xparent = "keyat9"
    If glbGP Then
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then    'Ticket #16189
            addNode xparent, "Database Setup", "gpDataSetup", "payweb"
            If glbCompSerial = "S/N - 2443W" Then 'Walters Ticket #25685 Franks 07/02/2014
                addNode xparent, "Princeton Time Database Setup", "gpPTimeDataSetup", "payweb"
            End If
        End If
        'If glbCompSerial = "S/N - 2259W" Then
        '    gsGPHold = False ' Oxford doesn't want the holding file yet.
        '    'addNode xparent, "Holding File", "gpHolding", "payweb"
        'End If
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then    'Ticket #16189
            addNode xparent, "Integration Setup", "gpIntgSetup", "payweb"
        End If
        'addNode xparent, "Export Attendance", "gpExpAttendance", "payweb"
        ''addNode xparent, "Reset Upload Flag", "gpResetUpload", "payweb"
        'addNode xparent, "Employee Master Transfer Report", "gpEmpMasterReport", "payweb"
        
        addNode xparent, "Income Code Matrix", "gpIncomeCodeMatrix", "payweb"
        If Not glbCompSerial = "S/N - 2259W" Then
            'addNode xparent, "Import Attendance", "gpImportAttendance", "payweb"
            'Ticket #17711 Frank 12/09/2009
            addNode xparent, "Import", "gpImportAttendance", "payweb"
            xparent = "gpImportAttendance"
            addNode xparent, "Attendance", "gpImportAtt_Att", "payweb"
            addNode xparent, "Entitlements/Attendance", "gpImportAtt_EntAtt", "payweb"
            addNode xparent, "Initial Data Load", "gpImportAtt_IDL", "payweb"
        End If
    End If
        
xparent = "keyat10"
    If glbMediPay Then
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then    'Ticket #16189
            addNode xparent, "Database Setup", "mpDataSetup", "payweb"
            addNode xparent, "Integration Setup", "mpIntgSetup", "payweb"
        End If
        'addNode xparent, "Import Attendance", "mpImpAttendance", "payweb"
        If gSec_Export_Employee And gSec_Export_Attendance And gSec_Import_Attendance Then  'Ticket #16189
            addNode xparent, "Export Attendance", "mpExpAttendance", "payweb"
                
            If glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab - Ticket #15094
                addNode xparent, "Reset Upload Flag", "mpResetUpload", "payweb"
            End If
        End If
        'addNode xparent, "Employee Master Transfer Report", "mpEmpMasterReport", "payweb"
    End If
    
'Simona - Leeds Grenville CAS ticket #14890
xparent = "keyat11"
    If glbCwis Then
        If gSec_Export_Employee And gSec_Export_Salaries And gSec_Export_Attendance Then    'Ticket #16189
            addNode xparent, "Database Setup", "cwDataSetup", "payweb"
        End If
    End If
        
'xparent = "keyat11"
'If glbWFC Then
'    addNode xparent, "Database Setup", "BonusDataSetup", "payweb"
'    addNode xparent, "Integration Setup", "BonusIntgSetup", "payweb"
'End If
    
'Ticket #24184 Franks 09/11/2013 - begin
xparent = "keyat12"
'If glbWFC Then
If glbWFC Or glbCompSerial = "S/N - 9999W" Then 'Ticket #25522 Franks 05/23/2014 - add 9999
    addNode xparent, "Maintenance", "keyat12M", "applicants"
    addNode xparent, "Reports", "keyat12R", "reports"
    addNode xparent, "Setup", "keyat12S", "setup"
    xparent = "keyat12M"
    addNode xparent, "Download File from FTP", "keyat12M1", "applicants"
    xparent = "keyat12R"
    addNode xparent, "XML Working Table Report", "keyat12R1", "reports"
    xparent = "keyat12S"
    addNode xparent, "FTP Setup", "keyat12S1", "setup"
    addNode xparent, "XML File Location", "keyat12S2", "setup"
End If
'Ticket #24184 Franks 09/11/2013 - end
    
'Ticket #26912 Franks 06/22/2015 - begin
xparent = "keyat15"
If glbCompSerial = "S/N - 2379W" Then
    addNode xparent, "Maintenance", "keyat15M", "applicants"
    addNode xparent, "Setup", "keyat15S", "setup"
    xparent = "keyat15M"
    addNode xparent, "Download File from FTP", "keyat15M1", "applicants"
    addNode xparent, "Upload File to FTP", "keyat15M2", "applicants"
    xparent = "keyat15S"
    addNode xparent, "FTP Setup", "keyat15S1", "setup"
End If
'Ticket #26912 Franks 06/22/2015 - end

'Ticket #29012 Franks 08/29/2016 - begin
If glbWFC And glbWFC_IncentivePlanFlag Then
    If Not glbtermopen Then
        xparent = "keyIP1"
        If gSec_WFC_IPExchangeRate Then ' gSec_Upd_Company Then
            addNode xparent, "Import Currency Exchange Table", "keyIP11", "applicants"
        End If
        If gSec_WFC_IPExchangeRate Then ' gSec_Upd_Company Then
            addNode xparent, "Currency Exchange Table", "keyIP12", "applicants"
        End If
    End If
End If

'Ticket #29012 Franks 08/29/2016 - end

tvwTree.Nodes(2).Expanded = True
'tvwTree.Nodes(3).Expanded = True 'Ticket #12828
tvwTree_NodeClick tvwTree.Nodes("root")
tvwTree.Nodes(2).EnsureVisible

If Not glbtermopen Then
    mnu_F_Employee.Enabled = True
    MainToolBar.ButtonS("find").ButtonMenus(1).Enabled = True

    'Ticket #25090 Franks 02/19/2014
    MainToolBar.ButtonS("find").ButtonMenus(2).Visible = gSec_Inq_Terminations
    
    'Ticket #22682 - Release 8.0
    MainToolBar.ButtonS("NewEmployee").Enabled = gSec_Add_NewHire And gSec_Inq_Basic And gSec_Upd_Basic
    
    'Ticket #29660 - New Contract Employee option is available only if you have right to add a New Hire
    If glbWFC Then
        MainToolBar.ButtonS("NewContractEmp").Enabled = gSec_Add_NewHire And gSec_Inq_Basic And gSec_Upd_Basic
    End If
    
Else
    mnu_F_Employee.Enabled = False
    MainToolBar.ButtonS("find").ButtonMenus(1).Enabled = False
    'MainToolBar.ButtonS("NewEmployee").Enabled = False
End If


MDIMain.panHelp(4).Caption = "Login User: " & glbUserID

Exit Sub
err_TreeSetting:
    If Err.Number = 5 Then Resume Next
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "TreeSetting", "SELECT")
    Resume Next
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub addNode(zParent As String, zName As String, Optional zKey As String, Optional zIconKey As String)
On Error GoTo err_addNode
    
    If zKey = "" Then zKey = LCase(zName)
    If zIconKey = "" Then
        If zParent = "root" Then
            zIconKey = zKey
        Else
            zIconKey = zParent
        End If
    End If
    tvwTree.Nodes.Add zParent, tvwChild, zKey, zName, zIconKey
    If zKey = "tkey15" Then 'for switch to active employees
        tvwTree.Nodes(zKey).Bold = True
        tvwTree.Nodes(zKey).ForeColor = RED ' &HFF&
    End If
    If zKey = "key113" Then
        tvwTree.Nodes(zKey).Bold = True
        tvwTree.Nodes(zKey).ForeColor = GREY
    End If
    If zKey = "key59c" Then
        tvwTree.Nodes(zKey).Bold = True
        tvwTree.Nodes(zKey).ForeColor = GREY
    End If
    'Ticket #18513 - Quick Setup for ESS - Disable the menu item
    If zKey = "key595" Then
        tvwTree.Nodes(zKey).ForeColor = vbGrayText
    End If
    If glbCompSerial = "S/N - 2172W" Then   'Lanark
        'Ticket #18518, Lanark not use Vacation/Sick Entitlement Master since they use GP
        If zKey = "key392" Then
            tvwTree.Nodes(zKey).ForeColor = vbGrayText
        End If
        If zKey = "key393" Then
            tvwTree.Nodes(zKey).ForeColor = vbGrayText
        End If
    End If
    If Not glbIsUseIHRDS Then 'Ticket #20310 Franks 05/10/2011
        If zKey = "MultipleDS" Then
            tvwTree.Nodes(zKey).ForeColor = vbGrayText
        End If
    End If
Exit Sub
err_addNode:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LOAD", "TreeSetting", zKey & " - " & zName) '"SELECT")
    Resume Next
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub tvwTree_Click()
    If Not Me.ActiveForm Is Nothing Then
        If Me.ActiveForm.name = "frmEEBASIC" Then
            If Me.ActiveForm.ChangeAction = NewRecord Then
                If Not isUpdated(Me.ActiveForm) Then Exit Sub
            End If
        End If
    Else
        '7.9 - Picture
        lstPanel.Visible = True
        lstView.Visible = True
    End If
End Sub

Private Sub tvwTree_GotFocus()
    If Me.ActiveForm Is Nothing Then
        set_Buttons
        
        '7.9 - Picture
        lstPanel.Visible = True 'False
        lstView.Visible = True  'False
    Else
    '    If Me.ActiveForm.name = "frmEEBASIC" Then
    '        If Me.ActiveForm.ChangeAction = NewRecord Then
    '            If Not isUpdated(Me.ActiveForm) Then Exit Sub
    '        End If
    '    End If
    End If
End Sub

Private Sub tvwTree_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
    Dim Reset As Boolean
    
    If frmFind Then
        frmFind = False
        Exit Sub
    End If
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(Str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If

    Set lstView.Icons = imlLarge
    
    If Node.Children = 0 Then
        If Not isUpdated(Me.ActiveForm) Then
            Exit Sub
        End If
        Call HandleCommand(Node.Key)
        lstView.ListItems.Clear
    End If
    
    Dim img
    Dim I
    If Node.Children = 0 Then
        If MDIMain.ActiveForm.Visible Then
            lstPanel.Visible = False
            lstView.Visible = False
        Else
            lstPanel.Visible = True 'False
            lstView.Visible = True  'False
        End If
    Else
        If Not MDIMain.ActiveForm.Visible Then
             lstPanel.Visible = True
             lstView.Visible = True
        End If
        If Not FirstTime Then
            lstView.ListItems.Clear
            'lstView.PictureAlignment = 5
            For I = 1 To tvwTree.Nodes.count
                If Not tvwTree.Nodes(I).Parent Is Nothing Then
                    If tvwTree.Nodes(I).Parent.Key = Node.Key Then
                        'lstView.PictureAlignment = 5
                        If tvwTree.Nodes(I).Children = 0 Or tvwTree.Nodes(I).Parent = "info:HR" Then
                            img = tvwTree.Nodes(I).Image
                        Else
                            img = "folder"
                        End If
                        'lstView.ListItems.Add , tvwTree.Nodes(I).Key, Left(tvwTree.Nodes(I).Text, 12), img
                        lstView.ListItems.Add , tvwTree.Nodes(I).Key, tvwTree.Nodes(I).Text, img
                        lstView.ListItems.Item(tvwTree.Nodes(I).Key).ToolTipText = tvwTree.Nodes(I).Text
                    End If
                End If
            Next I
        End If
    End If
End Sub

Private Sub lstView_DblClick()
    Dim ClickedKey As String
    'Dim I As Byte
    Dim I
    Dim img
    
    ClickedKey = lstView.SelectedItem.Key
    
    'if ther is no child form will be displayed.
    If tvwTree.Nodes(ClickedKey).Children = 0 Then
        Call HandleCommand(ClickedKey)
        lstView.ListItems.Clear
    Else
        lstView.ListItems.Clear
        tvwTree.Nodes(ClickedKey).Parent.Expanded = True
        tvwTree.Nodes(ClickedKey).Expanded = True
        For I = 1 To tvwTree.Nodes.count
            If Not tvwTree.Nodes(I).Parent Is Nothing Then
                If tvwTree.Nodes(I).Parent.Key = ClickedKey Then
                   'lstView.PictureAlignment = 5
                    If tvwTree.Nodes(I).Children = 0 Then
                        img = tvwTree.Nodes(I).Image
                    Else
                        img = "folder"
                    End If
                    'lstView.ListItems.Add , tvwTree.Nodes(I).Key, Left(tvwTree.Nodes(I).Text, 12), img
                    lstView.ListItems.Add , tvwTree.Nodes(I).Key, tvwTree.Nodes(I).Text, img
                    lstView.ListItems.Item(tvwTree.Nodes(I).Key).ToolTipText = tvwTree.Nodes(I).Text
                End If
            End If
        Next I
    End If
End Sub

Private Sub HandleCommand(CmdKey As String)

    Select Case CmdKey
        Case "Ke"
            Call mmnu_Find_Click
        Case "key6"
            Call mnu_Term_Inquiry_Click
        Case "KeyFindCandi"
            Call mmnu_FindCandi_Click
        Case "key6-1"
            Call mnu_Tran_Inquiry_Click
        Case "Key71"
            Call mmnu_EE_Basic_Click
        Case "key72"
            Call mmnu_EE_Status_Click
        Case "key73"
            Call mmnu_EE_Contact_Click
        Case "key74"
            Call mmnu_EE_Dependants_Click
        Case "key75"
            Call mmnu_EE_Payroll_Click
        Case "key76"
            Call mmnu_EE_EmpOther_Click
        Case "key77"
            Call mmnu_EE_EmpHistory_Click
        Case "key78"
            Call mmnu_EE_Flags_Click
        Case "key79"
            Call mmnu_EE_ADP_Click
        Case "Key81"
            Call mmnu_EE_Position_Click
        Case "key82"
            Call mmnu_EE_Salary_Click
        Case "key83"
            Call mmnu_EE_Performance_Click
        Case "key84"
            Call mmnu_EE_Other_Click
        Case "Key81a"
            Call mmnu_EE_Temp_CrossTrain_Position_Click
        Case "keyPayTrans"
            Call mmnu_EE_PayTrans_Click
        Case "key85"
            Call mmnu_EE_Benefits_Click
        Case "key86"
            Call mmnu_EE_ODollar_Click
        Case "key871"
            Call mmnu_SetPosition_Click
        Case "key872"
            Call mmnu_SetSalary_Click
        Case "key873"
            Call mmnu_SetPerformance_Click
        Case "key875"
            Call mmnu_ProfitSharing_Click
        Case "key88"
            Call mmnu_EE_Add_PayrollID_Data_Click
        Case "key90a"
            Call mmnu_EE_ASL_Click
        Case "Key91"
            Call mmnu_EE_Attendance_Click
        Case "key92"
            Call mmnu_Attend_History_Click
        Case "key93"
            Call mmnu_EE_VSE_Click
        Case "key94"
            Call mmnu_EE_VSO_Click
        Case "key95"
            Call mmnu_File_HrlyEntitlements_Click
        Case "key96"
            Call mmnu_EE_OvertimeO_Click
        Case "key97"
            Call mmnu_EE_WorkSchedule_Click
        Case "Key101"
            Call mmnu_EE_Associations_Click
        Case "key102"
            Call mmnu_EE_Seminars_Click
        Case "key103"
            Call mmnu_EE_FormalEd_Click
        Case "key106"
            Call mmnu_EE_Languages_Click
        Case "key104"
            Call mmnu_EE_Skills_Click
        Case "key105"
            Call mmnu_EE_Skills_Production_Click
        Case "key107"
            Call mmnu_EE_Succession_Click
        Case "Key108"
            Call mmnu_EE_UserDefine_Table_Click
        Case "Key109"
            Call mmnu_EE_Training_List_Click
        Case "Key111"
            Call mmnu_EE_EFollowup_Click
        Case "key112"
            Call mmnu_EE_vFollowup_Click
        Case "key113"
            'Call mmnu_EE_vWorkFlow_Click
        Case "Key121a"
            Call mmnu_EE_HS_W7_CompanyMaster_Click
        Case "Key121"
            Call mmnu_EE_HS_Incident_Data_Click
        Case "key122"
            Call mmnu_EE_HS_Injury_Click
        Case "key122a"
            Call mmnu_EE_HS_InjuryWF7_Click
        Case "key122b"
            Call mmnu_EE_HS_WSIBF9_Click     'Ticket #21463
        Case "key1221"
            Call mmnu_EE_HS_Reoccurrence_Click
        Case "key123"
            Call mmnu_EE_HS_Root_Cause_Click
        Case "key124"
            Call mmnu_EE_HS_Corrective_Click
        Case "key125"
            Call mmnu_EE_HS_Medical_Click
        Case "key126"
            Call mmnu_EE_HS_Contact_Click
        Case "key127"
            Call mmnu_EE_HS_Cost_Click
        Case "key128"
            Call mmnu_EE_HS_Acci_Cost_Click
        Case "key129"
            Call mmnu_EE_HS_Incident_Documents_Click
        Case "key130"
            Call mmnu_EE_HS_Company_Costs_Click
        Case "keyHSD1"
            Call mmnu_EE_HSDiv_Incident_Data_Click
        Case "keyHSD2"
            Call mmnu_EE_HSDiv_Injury_Click
        Case "keyHSD3"
            Call mmnu_EE_HSDiv_Root_Cause_Click
        Case "keyHSD4"
            Call mmnu_EE_HSDiv_Corrective_Click
        Case "keyHSD5"
            Call mmnu_EE_HSDiv_Medical_Click
        Case "keyHSD6"
            Call mmnu_EE_HSDiv_Contact_Click
        Case "keyHSD7"
            Call mmnu_EE_HSDiv_Incident_Documents_Click
        Case "key13"
            Call mmnu_EE_Counsel_Click
        Case "key14"
            Call mmnu_EE_Comments_Click
        Case "eesaldist"
            Call mmnu_EE_SalDist_Click
        Case "key14a"
            Call mmnu_EE_Cobra_Click
        Case "key151"
            Call mnu_Tlay_Click
        Case "key151-1"
            mmnu_Tran_Out_Click
        Case "key151-2"
            mmnu_Tran_In_Click
        Case "key151-3"
            mmnu_Tran_WFC_Div_Click
        Case "key152"
            Call mnu_EXTE_Click
        Case "key153"
            Call mnu_REAT_Click
        Case "key154"
            Call mmnu_Term_EE_Click
        Case "key155"
            Call mmnu_Term_Rehire_Click
        Case "key156a" '"key156"
            Call mmnu_Term_Retirement_Click
        Case "key156b"
            Call mmnu_Term_RetWorking_Click
        Case "key156c"
            Call mmnu_Term_RetRetiree_Click
        Case "key157"
            Call mmnu_Term_DeathProcess_Click
        Case "Key16a"
            Call mnu_F_Position_Click
        Case "Key16"
            Call mnu_Information_Click
        Case "Key16JB"
            Call mnu_FindJob_Click
        Case "Key16JC"
            Call mnu_JobMaster_Click
        Case "key17"
            Call mmnu_Pos_Skills_Click
        Case "key18"
            Call mnu_Pos_Eval_Click
        Case "key19"
            Call mnu_Pos_Course_Click
        Case "key191"
            Call mnu_Pos_Budget_Click
        Case "key192"
            Call mnu_Pos_PositionCtrl_Click
        Case "key19a"
            Call mmnu_Occ_Class_Click
        Case "keyPosDuties"
            Call mmnu_Pos_Duties_Click
        Case "keyPosResp"
            Call mmnu_Pos_Resp_Click
        Case "keyPosAppProc"
            Call mmnu_Pos_AppProc_Click
        Case "keyPosGrid"
            Call mmnu_Pos_Grid_Click
        Case "key19b"
            Call mmnu_Pos_BAND_Click
        Case "key193"
            Call mnu_Pos_DivDeptLnk_Click
        Case "accrualreport"
            Call mmnu_R_Accrual_Click
        Case "key201"
            Call mmnu_R_Attendance_Click
        Case "key202"
            Call mmnu_R_Attdpoint_Click
        Case "key203"
            Call mmnu_R_Attendance_Calendar_Click
        Case "key204"
            Call mmnu_R_AttdHist_Click
        Case "PersonalDayRpt"
            Call mmnu_R_AttdPersonalDayRpt_Click
        Case "key205"
            Call mmnu_R_Compensatory_Click
        Case "key2226_1"
            Call mmnu_R_FlexBank_Click
        Case "key206"
            Call mmnu_R_AttdCost_Click
        Case "Descrepancy"
            Call mmnu_R_AttWrkSchDescrepancy_Click
        Case "key207"
            Call mmnu_R_EEmergLeave_Click
        Case "key208"
            Call mmnu_R_Entitlements_Click
        Case "EnviroServ"
            Call mmnu_R_EnviroServ_Click
        Case "ESSReqTrnAud"
            Call mmnu_R_ESSReqTrnAudit_Click
        Case "key209"
            Call mmnu_R_HrEnt_Click
        Case "key205a"
            Call mmnu_R_OvertimeBank_Click
        Case "key205b"
            Call mmnu_R_OvertimeLostHours_Click
        Case "timesheet"
            Call mmnu_R_Timesheet_Click
        Case "timesheetWCost"
            Call mmnu_R_TimesheetWCost_Click
        Case "key210"
            Call mmnu_R_TimesheetStatus_Click
        Case "JournalEntry "
            Call mmnu_R_JournalEntry_Click
        Case "key211"
            Call mmnu_R_Birthday_Click
        Case "key212"
            Call mmnu_R_Dependents_Click
        Case "key213"
            Call mmnu_R_Email_Click
        Case "key214"
            Call mmnu_R_EEMaster_Click
        Case "key2141"
            Call mmnu_R_LOA_Click
        Case "key215"
            Call mmnu_R_Emergency_Click
        Case "key216"
            Call mmnu_R_EELabels_Click
        Case "key217"
            Call mmnu_R_Turnover_Click
        Case "key218"
            Call mmnu_R_EEProfile_Click
        Case "key220"
            Call mmnu_R_EEFlags_Click
        Case "key811"
            Call mmnu_R_EEGLDistribution_Click
        Case "key2343_1"
            Call mmnu_R_EESN2343_Click
        Case "key219"
            Call mmnu_R_EEPosition_Click
        Case "key2110"
            Call mmnu_R_Home_Click
        Case "key812"
            Call mmnu_R_EmployeeDates_Click
        Case "key813"
            Call mmnu_R_LengthOfService_Click
        'Hemu - 06/02/2004 Begin
        'Case "key2111"
        Case "key2114"
            Call mmnu_R_EEHistory_Click
        Case "key32a1"
            Call mmnu_R_PoPage_Click
        Case "key32a2"
            Call mmnu_R_StaffRatios_Click
        Case "key32a3"
            Call mmnu_R_WCLostTimeIncRate_Click
        Case "key32a4"
            Call mmnu_R_WCLostWrkHrRate_Click
        Case "key32a5"
            Call mmnu_R_ExternalHire_Click
        Case "key32a6"
            Call mmnu_R_InternalHire_Click
        Case "key32a7"
            Call mmnu_R_TurnoverRates_Click
        Case "key32a8"
            Call mmnu_R_PaidSickHr_Click
        Case "key2111a"
            Call mmnu_R_SIN_Click
        Case "key2112"
            Call mmnu_R_Tele_Ext_Click
        Case "key2113"
            Call mmnu_R_PlanEstablishment_Click
        Case "key22"
            Call mmnu_R_Counsel_Click
        Case "key22-L1"
            Call mmnu_R_DoorAccess_Click
        Case "key222"
            Call mmnu_R_DocumentType_Click
        Case "key231"
            Call mmnu_R_Associations_Click
        Case "key232"
            Call mmnu_R_Education_Click
        Case "key233"
            Call mmnu_R_Formal_Education_Click
        Case "key234"
            Call mmnu_R_Languages_Click
        Case "key235"
            Call mmnu_R_Skills_Click
        Case "key236"
            Call mmnu_R_Train_Matrix_Click
        Case "key23a"
            Call mmnu_R_Train_Plan_Click
        Case "key237"
            Call mmnu_R_Succession_Click
        Case "key238"
            Call mmnu_R_GapAnalysis_Click
        Case "key24"
            Call mmnu_R_Followup_Click
        Case "key24a"
            Call mmnu_R_FollowupEmailLog_Click
        Case "key801"
            Call mmnu_R_IWantToKnowYou_Click
        Case "key802"
            Call mmnu_R_ITHire_Click
        Case "key803"
            Call mmnu_R_ITNoticeOfChange_Click
        Case "key804"
            Call mmnu_R_NoticeOfChange_Click
        Case "key805"
            Call mmnu_R_PerfImproveActionPlan_Click
        Case "key806"
            Call mmnu_R_PerformanceReviewForm_Click
        Case "key807"
            Call mmnu_R_Separation_Click
        Case "key808"
            Call mmnu_R_TerminationForm_Click
        Case "key809"
            Call mmnu_R_UpdateMeeting_Click
        Case "key810"
            Call mmnu_R_Warning_Click
        Case "key251"
            Call mmnu_R_I_Body_Click
        Case "key252"
            Call mmnu_R_I_Day_Click
        Case "key253"
            Call mmnu_R_I_EE_Click
        Case "key254"
            Call mmnu_R_I_Trends_Click
        Case "key255"
            Call mmnu_R_I_Experience_Click
        Case "key256"
            Call mmnu_R_I_Incident_Click
        Case "key257"
            Call mmnu_R_I_Code_Click
        Case "key258"
            Call mmnu_R_I_Plant_Click
        Case "key259"
            Call mmnu_R_I_Shift_Click
        Case "key2510"
            Call mmnu_R_C_WCBI_Click
        Case "key2513"
            Call mmnu_R_C_CompanyAssoc_Cost_Click
'        Case "key2511" ' removed by Bryan 02/Dec/05
'            Call mmnu_R_C_Total_Click
        Case "key2512"
            Call mmnu_R_AccCost_Click
        Case "key27"
            Call mmnu_R_Position_Click
        Case "key28"
            Call mmnu_R_Seniority_Click
        Case "key291"
            Call mmnu_R_Password_Click
        Case "key292"
            Call mmnu_R_Table_Master_Click
        Case "key293"
            Call mmnu_R_CustomReport_Click
        Case "key293a"
            Call mmnu_R_HCASCustomReport_Click
        Case "key30"
            Call mmnu_R_Terminations_Click
        Case "key70"
            Call mmnu_R_SalVacIncr_Click
        Case "key99"
            Call mmnu_R_WorkSchedule_Click
        Case "key311a"
            'Ticket #16189 - Commented out previous logic for screen load security since now
            'only screens the user has access to will be visible.
            'If gSec_Rpt_Master_Benefits Then
                If glbLinamar Then
                    MsgBox "This function is not available."
                Else
                    Load frmRBenGroup
                    frmRBenGroup.ZOrder 0
                End If
            'Else
            '    MsgBox "You Do Not Have Authority For This Transaction"
            'End If
        Case "key311"
            Call mmnu_R_Benefit_Click
        Case "key311m"
            Call mmnu_R_BudPos_Click
        Case "key312m"
            Call mmnu_R_PosTablesExp_Click
        Case "key312"
            Call mmnu_R_CostER_Click
        Case "key313"
            Call mmnu_R_DolEnt_Click
        Case "key314"
            Call mmnu_R_OtherEarn_Click
        Case "keyPayTranRpt"
            Call mmnu_R_PayTrans_Click
        Case "keyProfitSharingRpt"
            Call mmnu_R_ProfitSharing_Click
        Case "keyRedCircledRpt"
            Call mmnu_R_RedCircled_Click
        Case "key315"
            Call mmnu_R_Salary_Click
        Case "key316"
            Call mmnu_R_Salary_Performance_Click
        Case "key317"
            Call mmnu_R_PerformanceReview_Click
        Case "key318"
            Call mmnu_R_Temp_CrossTraining_Click
        Case "key3F1"
            Call mmnu_R_AttendaceSignIn_Click
        Case "key3F2"
            Call mmnu_R_ATTDiscipline_Click
        Case "key3F3"
            Call mmnu_R_COCDiscipline_Click
        Case "key32"
            Call mmnu_Mass_AttHis_Click
        Case "key33"
                Call mmnu_Mass_Attendance_Click
        Case "key34"
            Call mmnu_Mass_Benefits_Click
        Case "key35"
            Call mmnu_Mass_Code_Click
        Case "key36"
            Call mmnu_mass_EducSemin_Click
        Case "key37"
            Call mmnu_Mass_CDE_Click
        Case "dooraccess"
            Call mmnu_Mass_DoorAccess_Click
        Case "key38"
            Call mmnu_Mass_Num_Change_Click
        Case "key38a"
            Call mmnu_Mass_term_Num_Change_Click
        Case "EnterLeave"
                Call mmnu_Mass_EnterLeave_Click
        Case "key40"
            Call mmnu_Mass_Followup_Click
        Case "key40-1"
            Call mmnu_IMP_Photo_Click
        Case "key41"
            Call mmnu_Mass_COE_Click
        Case "key42"
            Call mmnu_Mass_Position_Click
        Case "key42a"
            Call mmnu_Mass_EmployeePosition_Click
        Case "key43"
            Call mmnu_Mass_ReportAuth_Click
        Case "key44"
            Call mmnu_Mass_Salary_Click
        Case "key45"
            Call mmnu_Mass_TD1Dollar_Click
        Case "key46"
            Call mmnu_Mass_Terminations_Click
        Case "key98"
            Call mmnu_Mass_WorkSchedule_Click
        Case "key47d1"
            Call mmnu_Mass_BTI_QuarterEnd_Click
        Case "key47d2"
            Call mmnu_Mass_BTI_YTDCarryover_Click
        Case "key47d3"
            Call mmnu_Mass_BTI_YTDReduction_BD_Click
        Case "key47d4"
            Call mmnu_Mass_BTI_YTDReduction_NonBD_Click
        Case "key397"
            Call mmnu_Clear_Accrual_Click
        Case "key391"
            Call mmnu_holiday_Click
        Case "key392"
            If glbCompSerial = "S/N - 2172W" Then   'Lanark
                'Ticket #18518, Lanark not use Vacation Entitlement Master
            Else
                Call mmnu_VacEnt_Click
            End If
        Case "key392a"
            'Ticket #26154 - Oshawa Public Libraries - Vacation based on Seniority Hours
            Call mmnu_VacEarnedHours_Click
        Case "key392b"
            Call mmnu_VacEntDaily_Click
        Case "key392c"
            Call mmnu_VacEntDailySkippedLog_Click
        Case "key392d"
            Call mmnu_VacDailyAccDetails_Click
        Case "key393"
            If glbCompSerial = "S/N - 2172W" Then   'Lanark
            'Ticket #18518, Lanark not use Sick Entitlement Master
            Else
            Call mmnu_SickEnt_Click
            End If
        Case "key393a"
            Call mmnu_CurrentAccrYearEnd_Click
        Case "key393b"
            Call mmnu_AnnVacEnt_Click
        Case "key394"
            Call mmnu_HrsEnt_Click
        Case "key340"
            Call mmnu_HrsBasedEnt_Click
        Case "key395"
            Call mmnu_ZeroOutEnt_Click  'Hemu - 08/13/2003
        Case "key396"
            Call mmnu_RollOverEnt_Click  'Hemu - 08/13/2003
        
        Case "key395a"
            Call mmnu_ZeroOutHrEnt_Click  'Ticket #17924
        Case "key396a"
            Call mmnu_RollOverHrEnt_Click  ''Ticket #17924
            
        Case "key398"
            Call mmnu_VacPayPercentage_Click    'Ticket #25943 - Vacation Pay % for Hours Based Vacation Entitlement
        Case "key399"
            Call mmnu_HoursVacEntMst_Click    'Ticket #25943 - Hours Based Vacation Entitlement
        
        Case "key47"
            Call mmnu_File_Audit_Click
        Case "key47e"
            Call mmnu_Attendance_Audit_Click
        Case "key47ab"
            Call GroupBenefits
        Case "key47ac"
            Call BenefitCost
        Case "key47ad"
            Call mmnu_BonusDepartment_Master_Click
        Case "key47af"
            Call GroupBenefitMatrix
        Case "key47ag"
            Call GroupManulifeRule
        Case "key47ah"
            Call GroupOMERS_Formula
        Case "key47ai"
            Call mmnu_BenefitRates
        Case "key48"
            Call mmnu_Company_Master_Click
        'Case "key48C"
        '    Call mmnu_Company_Preference_Click
        Case "key48C_General"
            Call mmnu_General_Click
        Case "key48C_EmailNotification"
            Call mmnu_Email_Notification_Click
        Case "key48C_FileLocation"
            Call mmnu_FileLocation_Click
        Case "key48B"
            Call mmnu_Counsel_Absence_Click
        Case "key48D"
            Call mmnu_CourseCode_Master_Click
        'Case "key48B2"
        '    Call mmnu_Counsel_LE_Click
        Case "key49"
            Call mmnu_CustomReport_Master_Click
        Case "key50"
            Call mmnu_Department_Master_Click
        Case "key50a"
            Call mmnu_OHRSDepartment_Master_Click
        Case "key51"
            Call mmnu_Division_Master_Click
        Case "SalDist"
            Call mmnu_SalDist_Master_Click
        Case "paycategory"
            Call mmnu_PayCategory_Master_Click
        Case "paybenmatrix"
            Call mmnu_PayMatrixBenefit_Click
        Case "ChargeCode"
            Call mmnu_ChargeCode_Master_Click
        Case "ProjectCode"
            Call mmnu_ProjectCode_Master_Click
        Case "key58b"
            Call mmnu_AttendCode_Matrix_Click
        Case "key58d"
            Call mmnu_FollowUpCodeEmail_Matrix_Click
        Case "key58e"
            Call mmnu_DeptGL_Matrix_Click
        Case "Machine"
            Call mmnu_Machine_Master_Click
        Case "MultipleDS"
            Call mmnu_Multiple_Data_Source_Click
        Case "key510"
            Call mmnu_Disciplinary_Steps_Click
        Case "key5help1"
            Call mmnu_Help_Desc_Click
        Case "key52"
            Call mmnu_Lgr_Master_Click
        Case "key53"
            Call mmnu_Opus_Payroll_Click
        Case "key54"
            Call mmnu_Label_Click
        Case "key54_Lbl1"
            Call mmnu_Label1_Click
        Case "key54_Lbl2"
            Call mmnu_Label2_Click
        Case "key54_Lbl3"
            Call mmnu_Label3_Click
        Case "key54_Lbl4"
            Call mmnu_Label4_Click
        Case "key54_Lbl5"
            Call mmnu_Label5_Click
        Case "key54_Lbl6"
            Call mmnu_Label6_Click
        Case "key54_Lbl20"
            Call mmnu_Label20_Click
        Case "key54_Lbl7"
            Call mmnu_Label7_Click
        Case "key54_Lbl8"
            Call mmnu_Label8_Click
        Case "key54_Lbl9"
            Call mmnu_Label9_Click
        Case "key54_Lbl10"
            Call mmnu_Label10_Click
        Case "key54_Lbl11"
            Call mmnu_Label11_Click
        Case "key54_Lbl12"
            Call mmnu_Label12_Click
        Case "key54_Lbl13"
            Call mmnu_Label13_Click
        Case "key54_Lbl14"
            Call mmnu_Label14_Click
        Case "key54_Lbl15"
            Call mmnu_Label15_Click
        Case "key54_Lbl16"
            Call mmnu_Label16_Click
        Case "key54_Lbl21"
            Call mmnu_Label21_Click
        Case "key54_Lbl17"
            Call mmnu_Label17_Click
        Case "key54_Lbl18"
            Call mmnu_Label18_Click
        Case "key54_Lbl19"
            Call mmnu_Label19_Click
        Case "key550"
            Call mmnu_MarketLine_Click
        Case "key55"
            Call mmnu_New_Hire_Click
        Case "key56"
            Call mmnu_File_Payroll_Matrix_Click
        Case "key56-0"
            Call mmnu_File_Pay_Pariod_Master_Click
        Case "key56-1"
            Call mmnu_Close_Pay_Period_Click
        Case "keyLinamarCode-BNCD"
            Call mnu_Benefit_Click
        Case "keyLinamarCode-HMOP"
            Call mnu_Home_Operation_Click
        Case "keyLinamarCode-HMLN"
            Call mnu_Home_Line_Click
        Case "keyLinamarCode-HMWC"
            Call mnu_Home_Work_Click
        Case "keyLinamarCode-HMSF"
            Call mnu_Home_Shift_Click
        Case "keyLinamarCode-EDSE"
            Call mnu_Operation_Click
        Case "keyLinamarCode-EDRG"
            Call mnu_ProductLine_Click
        Case "keyLinamarCode-SHFT"
            Call mnu_Shift_Click
        Case "keyLinamarCode-EDSK"
            Call mnu_Skill_code_Click
        Case "keyLinamarCode-RGSE"
            Call mnu_ProductLine_Operation_Click
        Case "key57"
            Call mmnu_Prov_Master_Click
        Case "key5171"
            Call mmnu_Root_Cause_Event_Click
        Case "key5172"
            Call mmnu_Root_Cause_Immediate_Click
        Case "key5173"
            Call mmnu_Root_Basic_Underlying_Click
        Case "key58"
            Call mmnu_Table_Master_Click
        Case "key58c"
            Call mmnu_Table_Master_CodeLinks_Click
        Case "key58a"
            Call mmnu_Table_Attendance_Group_Master_Click
        Case "key591"
            Call mnu_File_ChgPass_Click
        Case "key592"
            Call mnu_File_Secure_Click
        Case "key592-L1"
            Call mnu_File_Door_Click
        Case "key592-L2"
            Call mnu_File_DoorName_Click
        Case "key593"
            Call mnu_File_EmailSetup_Click
        Case "key594"
            Call mnu_File_EmailLog_Click
        Case "WorkSchRule"
            Call mnu_WorkScheduleRule_Click
        Case "DashboardSetup"
            Call mnu_DashboardSetup_Click
        Case "OnCallHrs"
            Call mnu_OnCallHours_Click
        Case "key595"
            'Ticket #18513 - Disable the menu item
            'Call mnu_File_QuickSetupESS_Click
        Case "key59a"
            Call mnuUnionSickBank_Click
        Case "key59b"
            Call mnuTerminationCauseLink_Click
        Case "key59c"
            'Call mnuWorkFlowMaster_Click
        Case "key60" 'added by Bryan 12/07/05 Ticket #8922
            Call mnuManpower_Click
        Case "key61"
            Call mmnu_OvertimeMaster_Click
        Case "key62"
            Call mmnu_SalaryIncr_Click
        Case "key63"
            Call mmnu_VacationIncr_Click
        Case "key64"
            Call mmnu_PosGrp_PerfCat_Click
        Case "key65"
            Call mmnu_Employee_Type_Matrix_Click
        Case "key66"
            Call mmnu_Person_Completing_Form7_Click
        Case "tkey15"
            Call mmnu_Active_Click
        Case "keypw631"
            Call mmnu_PayWeb_Setup_Click
        Case "keypw632"
            Call mmnu_PayWeb_Code_Matrix_Click
        Case "keypw611"
            Call mmnu_PayWeb_Exp_Attd_Click
        Case "keypw612"
            Call mmnu_PayWeb_Exp_IDL_Click
        Case "keypw613"
            Call mmnu_PayWeb_Exp_Ongoing_Click
        Case "keypw614"
            Call mnu_Payweb_Reset_Click
        Case "keypw621"
            Call mmnu_PayWeb_Imp_Attd_Click
        Case "keypw622"
            Call mmnu_PayWeb_Imp_IDL_Click
        Case "keypw623"
            Call mmnu_PayWeb_Imp_YTD_Click
        Case "keyvd70"
            Call mmnu_Vadim_Pay_Code_Click
        Case "keyvd71"
            Call mmnu_Vadim_Setup_Click
        Case "keyvd72"
            Call mmnu_Vadim_Database_Setup_Click
        Case "keyvd73"
            Call mmnu_Vadim_Report_Click
        Case "keyvd74"
            Call mmnu_Vadim_Accrual_Class_Click
        Case "keyvd75"
            Call mmnu_Vadim_Code_Sync_Click
        Case "keyvd75a"
            Call mmnu_Vadim_IHRCode_Sync_Click
        Case "keyvd76"
            Call mmnu_Vadim_Import_Table_Click
        Case "keyvd76a"
            Call mmnu_Code_Matrix_Click
        Case "keyvd77"
            Call mmnu_Vadim_Att_Sync_Click
        Case "keyvd78"
            Call mmnu_Vadim_IDL_Accural_Click
        Case "keyvd79"
            Call mmnu_Vadim_Salary_Report_Click
        Case "advDataSetup"
            Call mmnu_Other_Data_Setup_Click("Advanced Tracker")
        Case "advDeleteAttnd"
            Call mmnu_Integration_Click("Advanced Tracker", "Delete Attendance Records")
        Case "advIntgSetup"
            Call mmnu_Integraion_Setup_Click("Advanced Tracker")
        Case "advSelectionSetup"
            Call mmnu_Integration_Selection_Click("Advanced Tracker")
        Case "advImpAttendance" '
            Call mmnu_Integration_Click("Advanced Tracker", "Import Attendance")
        Case "advBankSync"
            Call mmnu_Time_Bank_Sync_Click("Advanced Tracker")
        Case "advResetUpload"
            Call mmnu_Integration_Click("Advanced Tracker", "Reset Upload Flag")
        Case "advExpEmpTblMst_IDL"
            Call mmnu_Adv_ExpEmpTblIDL_Click("Advanced Tracker")
        Case "advExpVacSickBalance"
            Call mmnu_Adv_ExpVacSickBalance_Click("Advanced Tracker")
        Case "gpDataSetup"
            Call mmnu_Other_Data_Setup_Click("Great Plains")
        Case "gpPTimeDataSetup" 'Ticket #25685 Franks 07/02/2014
            Call mmnu_Other_Data_Setup_Click("Princeton Time")
        Case "gpIntgSetup"
            Call mmnu_Integraion_Setup_Click("Great Plains")
        Case "gpHolding"
            Call mmnu_GP_HOLDING_Click("Great Plains")
        Case "gpIncomeCodeMatrix"
            Call mmnu_GP_IncomeCodeMatrix_Click("Great Plains")
        Case "gpImportAtt_Att" '"gpImportAttendance"
            Call mmnu_GP_ImportAttendance_Click("Great Plains")
        Case "gpImportAtt_EntAtt"
            Call mmnu_GP_ImportEntAtt_Click("Great Plains")
        Case "gpImportAtt_IDL"
            Call mmnu_GP_ImportAttIDL_Click("Great Plains")
        Case "mpDataSetup"
            Call mmnu_Other_Data_Setup_Click("MediPay")
        Case "mpIntgSetup"
            Call mmnu_Integraion_Setup_Click("MediPay")
        Case "mpImpAttendance"
            Call mmnu_Integration_Click("MediPay", "Import Attendance")
        Case "mpExpAttendance"
            Call mmnu_Integration_Click("MediPay", "Export Attendance")
        Case "mpEmpMasterReport"
            Call mmnu_Integration_Click("MediPay", "Employee Master Transfer Report")
        Case "mpResetUpload"
            Call mmnu_Integration_Click("MediPay", "Reset Upload Flag")
        Case "BonusDataSetup"
            Call mmnu_Other_Data_Setup_Click("Bonus System")
        Case "BonusIntgSetup"
            Call mmnu_Integraion_Setup_Click("Bonus System")
        Case "key47b"
            Call submnu_USData_Click
        Case "key47c"
            Call submnu_Rep_OccupGroup_Click
        Case "key47g" 'Ticket #18790
            Call submnu_PurgeTermEEO_Click
        Case "key51a1"
            Call submnu_Plan_Data_Click
        Case "key51a2"
            Call submnu_Survey_Data_Click
        Case "key51a34"
            Call submnu_Rep_ComplWork_Click
        Case "key51a35"
            Call submnu_Rep_EmplStatus_Click
        Case "key51a36"
            Call submnu_Rep_Work_Click
        Case "key51a37"
            Call submnu_Rep_EmpEquityVitalAire_Click
        Case "BonusDataSetup"
            Call mmnu_Other_Data_Setup_Click("MediPay")
        Case "BonusIntgSetup"
            Call mmnu_Integraion_Setup_Click("MediPay")
        Case "key32a9"
            Call mnu_Manpower_Plan_Click 'added by Bryan 14/07/05 Ticket #8921
        Case "key32a10"
            Call mnu_DailyManpower_Click 'added by Bryan 9/Sep/05 Ticket #9235
        Case "gpPosting"
            Call mnu_GPPosting_Click 'added by George Feb 23,2006 Ticket #9965
        Case "tsShiftRpt"
            Call mnu_ShiftSchedule_Click 'added by Bryan 31/Oct/05 Ticket#9630
        Case "tsQuarter"
            Call mnu_QuarterlyReport_Click 'added by Bryan 07/Nov/05 Ticket#9720
        Case "tsHSSheet"
            Call mnu_HSWorkSheet_Click 'added by Bryan 09/Nov/05 Ticket#9720
        Case "tsCBSheet"
            Call mnu_CBWorksheet_Click 'added by Bryan 30/Nov/05 Ticket#9721
        Case "keyPPct"
            Call mnu_PensionPct_Click 'added by Bryan 28/Dec/05 Ticket#9771
        Case "EmpFlags"
            Call mmnu_EmpFlags_Click
        Case "KeyGLDist" 'added by Bryan 22/Feb/06 Ticket#10308
            Call mmnu_GLDIST_Click
        Case "PerfReview" 'Friesens Corporation Ticket#10844
            Call mmnu_PerformanceReview_Click
        Case "PerfReviewH" 'Friesens Corporation Ticket#10844
            Call mmnu_PerformanceReviewH_Click
        Case "EMLSETUP"
            Call mmnu_EmlSetup_Click
        Case "keypEI" 'Ticket #13142
            Call mmnu_ImportExport_Click
        Case "key239" 'User Defined Table Report
            Call mmnu_UserDefinedTableReport_Click
        Case "key240"   'Friesens - Ticket #16189
            Call mmnu_R_Req_Course_Hist_Click
        Case "key209a" 'Future Entitlement
            Call mmnu_FutureEntitlementReport_Click
        Case "cwDataSetup"
            Call mmnu_Other_Data_Setup_Click("CWIS") 'Simona - Leeds Grenville CAS - ticket#14890
        'Ticket #24184 Franks 09/11/2013 - begin
        Case "keyat12M1"
            Call mmnu_SF_Download_Click
        Case "keyat12R1"
            Call mmnu_SF_XMLRPT_Click
        Case "keyat12S1"
            Call mmnu_SF_FTPSETUP_Click
        Case "keyat12S2"
            Call mmnu_SF_XML_LOCATION_Click
        'Ticket #24184 Franks 09/11/2013 end
        'Ticket #26912 Franks 06/22/2015 - begin
        Case "keyat15M1"
            Call mmnu_Sys247_Download_Click
        Case "keyat15M2"
            Call mmnu_Sys247_Upload_Click
        Case "keyat15S1"
            Call mmnu_Sys247_FTPSETUP_Click
        'Ticket #26912 Franks 06/22/2015 end
        Case "key100"
            'Ticket #23409 - Samuel, Son & Co., Limited - Discipline Audit Table Report
            Call mmnu_File_CounselAudit_Click
        Case "key200"
            'Release 8.0 - Ticket #24361: Add Email Address import under Mass Updates menu
            Call mmnu_Mass_ImportEmailAddress_Click
        Case "key199a"
            'Release 8.1 - Ticket #27244: Import document Attachment under Mass Updates menu
            Call mmnu_Mass_ImportAttachment_Click
        Case "key199b"
            'Release 8.1 - Ticket #27244: Document Type Information update under Mass Updates menu
            Call mmnu_Mass_DocTypeInfoUpdate_Click
        Case "key110"
            'Macaulay - Ticket #25015 - Seniority Date Calculation
            Call mmnu_SeniorityDateCalculation_Click
        'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc. - begin
        Case "key5JobClass1"
            Call mmnu_JobFamily_Click
        Case "key5JobClass2"
            Call mmnu_SubJobFamily_Click
        Case "key5JobClass3"
            Call mmnu_GroupJob_Click
        'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc. - end
        
        'Ticket #29013 Franks 08/23/2016 - begin
        'WFC Incentive Plan
        Case "keyIP11"
            Call mmnu_IncentivePlan_ImpCurrency_Click
        Case "keyIP12"
            Call mmnu_IncentivePlan_CurrencyTable_Click
        Case "keyIP2"
            Call mmnu_IncentivePlan_Factors_Click
        Case "keyIP3"
            Call mmnu_IncentivePlan_CreateSpreadsheet_Click
        Case "keyIP4"
            Call mmnu_IncentivePlan_ImpSpreadsheet_Click
        Case "keyIP5"
            Call mmnu_IncentivePlan_UptOtherEarnings_Click
        Case "keyIP6"
            Call mmnu_IncentivePlan_PreparePayroll_Click
        Case "keyIP7"
            Call mmnu_IncentivePlan_PrintSpreadsheet_Click
        Case "keyIP8"
            Call mmnu_IncentivePlan_PrintEmpLetter_Click
        'Ticket #29013 Franks 08/23/2016 - end
        
        Case "keyAppT001"
            Call mmnu_AppTrack_LetterByPosType_Click
        Case "keyAppT002"
            Call mmnu_AppTrack_AppFormWorkflow_Click
        Case "keyAppT003"
            Call mmnu_AppTrack_AppFormDefaults_Click
            
    End Select
    
Call UnloadFrms("newform")
End Sub

Private Sub remNode()
    tvwTree.Nodes.Clear
    Call TreeSetting
End Sub

Private Sub MainToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RelateMode As RelateModeEnum
Dim xMenu As Menu, X

    Select Case Button.Key
    Case "close"
        Call mnu_F_Close_Click
    Case "preview"
        Call mnu_F_Preview_Click
    Case "print"
        Call mnu_F_Print_Click
    Case "NewEmployee"
        Call mnu_NewEmployee_Click
    Case "NewRecord"
        Call mnu_E_NewRecord_Click
    Case "save"
        Call mnu_E_Save_Click
    Case "cancel"
        Call mnu_E_Cancel_Click
    Case "delete"
        Call mnu_E_Delete_Click
    Case "up"
        Call mnu_M_Up_Click
    Case "down"
        Call mnu_M_Down_Click
    Case "find"
        RelateMode = get_RelateMode(Me.ActiveForm)
        Call clkFind(RelateMode)
    Case "massdelete"
        Call mnu_M_Delete_Click
    Case "massupdate"
        Call mnu_M_Update_Click
    Case "massadd"
        Call mnu_M_Add_Click
    Case "word"
        Call mnu_WORD_Click
    Case "excel"
        Call mnu_EXCEL_Click
    Case "mail"
        Call mnu_Mail_Click
    Case "help"
        Call DispHelp(Me.ActiveForm)
    Case "hrsoft" 'Ticket #24184 Franks 12/04/2013
        Call mnu_HRsoft_Click
    Case "NewContractEmp" 'Ticket #29660 - Contract Employees
        Call mnu_NewContractEmployee_Click
    End Select
End Sub

Private Sub MainToolBar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "findA"
        Call GET_EMP
    Case "findTerm"
        Call mnu_F_TEmployee_Click
    Case "findJOB" 'Ticket #28118 Franks 02/01/2016
        Call mnu_F_JobMaster_Click
    Case "findPOS"
        Call mnu_F_Position_Click
    Case "findDiv"
        Call mnu_F_Division_Click
    Case "findCandi" 'Ticket #25676 Franks 07/29/2014
        Call mnu_F_Candidate_Click
    End Select
End Sub

Private Sub clkNew(NMode As String)
Dim OLDid
On Error GoTo Err_Deal
    If Not isUpdated(Me.ActiveForm) Then Exit Sub

    If NMode = "NewRecord" Then
        'Call UnloadFrms("newrecord")
        Call Me.ActiveForm.cmdNew_Click
    ElseIf NMode = "ContractEmployee" And glbWFC Then
        'Ticket #29660 - Contract Employees - Adding a New Hire
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        glbWFCContractEmployee = True
        glbOnTop = ""   'So any screen open will be closed
        Call UnloadFrms
        frmEContEmpDemo.ChangeAction = NewRecord
        Call frmEContEmpDemo.cmdNew_Click
        If glbTrsEE_ID <> "" And glbLEE_ID <> 0 Then
            frmEContEmpDemo.Show 1
        End If
        
    Else
        glbCandidate = 0 'Ticket #24184 Franks 09/11/2013
        glbHRSoftType = ""
        glbWFCContractEmployee = False
        Call UnloadFrms
        frmEEBASIC.ChangeAction = NewRecord
        Call frmEEBASIC.cmdNew_Click
    End If
'Exit Sub
Err_Deal:
    If Err = 364 Then Resume Next
End Sub

Private Sub clkFind(RMode As RelateModeEnum)
Dim OLDid
Dim OJobSec
On Error GoTo Err_Deal
    
    If Not isUpdated(Me.ActiveForm) Then Exit Sub
    
    Select Case RMode
    Case RelatePos
        OLDid = glbPos
        OJobSec = glbJobSection
        If glbWFC Then 'Ticket #25911 Franks 10/07/2014
            frmJOBSWFC.Show 1
            If glbPos <> OLDid Or glbJobSection <> OJobSec Then GoTo DisplayForm
        Else
            frmJOBS.Show 1
            If glbPos <> OLDid Then GoTo DisplayForm
        End If
        'If glbPos <> OLDid Then GoTo DisplayForm
    Case RelateJobMaster 'Ticket #28118 Franks 02/01/2016
            'frmJOBSWFC.Show 1
            OLDid = glbJobMaster
            frmMJobMaster.Show 1
            If glbJobMaster <> OLDid Then GoTo DisplayForm
    Case Else
        OLDid = glbLEE_ID
        Call GET_EMP
    '    If glbLEE_ID <> OLDid Then GoTo DisplayForm
    End Select

Exit Sub
DisplayForm:
        Call ReDisplayForms(RMode)
Err_Deal:
    If Err = 364 Then Exit Sub
    If Err = 91 Then Resume Next
End Sub

Public Sub Calculate_DurhamCHCEnt() 'Ticket #27765 Franks 02/12/2016
    Dim rsVacEnt As New ADODB.Recordset
    Dim rsEntRunDt As New ADODB.Recordset
    Dim SQLQ
    Dim selSQLQ
    
    'If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'testing mode
    '    Exit Sub
    'End If
    
    SQLQ = "SELECT * FROM HR_ENTRUNDATE"
    rsEntRunDt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEntRunDt.EOF Then Exit Sub
    If Not rsEntRunDt.EOF Then
            ''run it once a day
            If CVDate(rsEntRunDt("EN_LAST_RUN_DATE")) = CVDate(Format(Now, "mmm/dd/yyyy")) Then Exit Sub
            ''run it once a month
            'If month(CVDate(rsEntRunDt("EN_LAST_RUN_DATE"))) = month(CVDate(Date)) And Year(rsEntRunDt("EN_LAST_RUN_DATE")) = Year(Date) Then Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = "Calculate Accrued to Date Vacation..."

    'Call procedure that will calculate the Pension Entitlement for Samuel
    'If PensionCalculation4Samuel = False Then GoTo Do_not_update
    If frmSVacEnt.Auto_AccruedVacEnt_Upd_DurhamCHC_Run = False Then GoTo Do_not_update
    
    If rsEntRunDt.EOF Then
        rsEntRunDt.AddNew
        rsEntRunDt("EN_COMPNO") = "001"
    End If
    'Update Last Run Date table - to indicate the entitlement calc was run today
    rsEntRunDt("EN_LAST_RUN_DATE") = Date 'CVDate(Format(Now, "mmm/dd/yyyy"))
    rsEntRunDt("EN_LUSER") = glbUserID
    rsEntRunDt("EN_LDATE") = Date
    rsEntRunDt("EN_LTIME") = Time$
    rsEntRunDt.Update
    rsEntRunDt.Close
    MDIMain.panHelp(1).Caption = "Accrued to Date Vacation update complete"
    MDIMain.panHelp(0).Caption = ""
    Screen.MousePointer = DEFAULT
    
Exit Sub

Do_not_update:
    rsEntRunDt.Close
    MDIMain.panHelp(1).Caption = "An error occurred in Accrued to Date Vacation update"
    MDIMain.panHelp(0).Caption = ""
    
End Sub

Public Sub Calculate_Entitlement() 'for Samuel - Ticket #20589
    Dim rsVacEnt As New ADODB.Recordset
    Dim rsEntRunDt As New ADODB.Recordset
    Dim SQLQ
    Dim selSQLQ
    
    If Not (glbCompSerial = "S/N - 2382W") Then
        Exit Sub
    End If
    SQLQ = "SELECT * FROM HR_ENTRUNDATE"
    rsEntRunDt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEntRunDt.EOF Then Exit Sub
    If Not rsEntRunDt.EOF Then
        If glbSamuel Then
            'run it once a month
            If month(CVDate(rsEntRunDt("EN_LAST_RUN_DATE"))) = month(CVDate(Date)) And Year(rsEntRunDt("EN_LAST_RUN_DATE")) = Year(Date) Then Exit Sub
            If Not IsNull(rsEntRunDt("EN_USERLIST")) Then
                If InStr(1, rsEntRunDt("EN_USERLIST"), glbUserID & "+") = 0 Then
                    Exit Sub
                End If
            End If
        Else
            If CVDate(rsEntRunDt("EN_LAST_RUN_DATE")) = CVDate(Format(Now, "mmm/dd/yyyy")) Then Exit Sub
        End If
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = "Updating Pension Entitlement..."

    'Call procedure that will calculate the Pension Entitlement for Samuel
    If PensionCalculation4Samuel = False Then GoTo Do_not_update
    
    'If rsEntRunDt.EOF Then rsEntRunDt.AddNew
    'Update Last Run Date table - to indicate the entitlement calc was run today
    rsEntRunDt("EN_LAST_RUN_DATE") = Date 'CVDate(Format(Now, "mmm/dd/yyyy"))
    rsEntRunDt("EN_LUSER") = glbUserID
    rsEntRunDt("EN_LDATE") = Date
    rsEntRunDt("EN_LTIME") = Time$
    rsEntRunDt.Update
    rsEntRunDt.Close
    MDIMain.panHelp(1).Caption = "Pension Entitlement update complete"
    MDIMain.panHelp(0).Caption = ""
    Screen.MousePointer = DEFAULT
    
Exit Sub

Do_not_update:
    rsEntRunDt.Close
    MDIMain.panHelp(1).Caption = "An error occurred in Pension Entitlement update"
    MDIMain.panHelp(0).Caption = ""
End Sub
Private Sub Calculate_Vacation_Entitlement()
    
    
    'For Mitchel Plastics - Ultra Manufacturing Ticket # 8104
    Dim rsVacEnt As New ADODB.Recordset
    Dim rsEntRunDt As New ADODB.Recordset
    Dim SQLQ
    Dim selSQLQ
    MDIMain.panHelp(0).Caption = "Updating Vacation Entitlement...Please wait"

    SQLQ = "SELECT LAST_RUN_DATE FROM ENTRUNDATE"
    rsEntRunDt.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    If Not rsEntRunDt.EOF Then
        If CVDate(rsEntRunDt("LAST_RUN_DATE")) = CVDate(Format(Now, "mmm/dd/yyyy")) Then Exit Sub
    End If
        
    'Call procedure that will calculate the Vacation Entitlement for Ultra Manufacturing
    If frmSVacEnt.Automatic_VacEntitlement_Update_Run = False Then GoTo Do_not_update

    If rsEntRunDt.EOF Then rsEntRunDt.AddNew
    'Update Last Run Date table - to indicate the entitlement calc was run today
    rsEntRunDt("LAST_RUN_DATE") = CVDate(Format(Now, "mmm/dd/yyyy"))
    rsEntRunDt.Update
    rsEntRunDt.Close
    MDIMain.panHelp(0).Caption = "Vacation Entitlement update complete"
    MDIMain.panHelp(0).Caption = ""
Exit Sub

Do_not_update:
    rsEntRunDt.Close
    MDIMain.panHelp(0).Caption = "An error occurred in Vacation Entitlement update"
    MDIMain.panHelp(0).Caption = ""
End Sub

Private Sub mnu_ShiftSchedule_Click()
    frmRShift.Show
End Sub

Private Sub mnu_QuarterlyReport_Click()
    frmRQuarter.Show
End Sub

Private Sub mnu_HSWorkSheet_Click()
    frmRExcelRpt.Caption = "Health & Safety report"
    frmRExcelRpt.Rptname = "sn2369HS"
    frmRExcelRpt.Show
    frmRExcelRpt.ZOrder 0
End Sub

Private Sub mnu_CBWorksheet_Click()
    frmRExcelRpt.Caption = "Bonus report"
    frmRExcelRpt.Rptname = "rzChalBonus"
    frmRExcelRpt.Show
    frmRExcelRpt.ZOrder 0
End Sub

Private Function DoLaunchWord(oWordApplication As Object) As Boolean
    On Local Error GoTo DoLaunchEH
    Set oWordApplication = CreateObject(mcsWordApplication)
    With oWordApplication
        .Visible = True
        .Activate
    End With
    DoLaunchWord = True
DoLaunchEH:
End Function

Private Function DoLaunchExcel(objExcelApplication As Object) As Boolean
    On Local Error GoTo DoLaunchEH
    Set objExcelApplication = CreateObject(mcsExcelApplication)
    With objExcelApplication
        .Visible = True
        .Activate
    End With
    'Set oExcelWorkbook = oExcelApplication.Workbooks.Add
    DoLaunchExcel = True
DoLaunchEH:
End Function

Private Sub mmnu_GLDIST_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Inq_GLDist Then
        frmEGLDist.Show
        frmEGLDist.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_UserDefinedTableReport_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gSec_Rpt_User_Defined_Table Then
        frmRUserDef.Show
        frmRUserDef.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub mmnu_FutureEntitlementReport_Click()
    'Ticket #16189 - Commented out previous logic for screen load security since now
    'only screens the user has access to will be visible.
    'If gsec_rpt_Future_Entitlement Then
        frmRNextEnt.Show
        frmRNextEnt.ZOrder 0
    'Else
    '    MsgBox "You Do Not Have Authority For This Transaction"
    'End If
End Sub

Private Sub DispHelp(frmName As Form)
    Dim rsHelp As New ADODB.Recordset
    Dim SQLQ As String
    Dim ConName As String
    Dim xTablName As String, xCode As String, xCodeCaption As String
    Dim xDispHelp As String
    
    On Error GoTo Line_Err
        
    If Me.ActiveForm Is Nothing Then Exit Sub
    ConName = frmName.ActiveControl.name
    If ConName = "clpCode" Then
        xTablName = frmName.ActiveControl.TablName
        xCode = frmName.ActiveControl.Text
        xCodeCaption = frmName.ActiveControl.Caption
        If Len(xTablName) > 0 And Len(xCode) > 0 Then
            SQLQ = "SELECT * FROM HRHELP WHERE HP_TABL_NAME = '" & xTablName & "' "
            SQLQ = SQLQ & "AND HP_TABL_KEY = '" & xCode & "' "
            rsHelp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsHelp.EOF Then
                xDispHelp = rsHelp("HP_MEMO")
            End If
            rsHelp.Close
            If Len(xDispHelp) > 0 Then
                frmDispHelp.lblCode = xCode & " " & xCodeCaption
                frmDispHelp.memComments = xDispHelp
                frmDispHelp.Show 1
            End If
        End If
    End If
Exit Sub
Line_Err:

End Sub

Public Function gSec_Inq_Master_Table_Exists(xKey As String) As Boolean
    On Error GoTo gSec_Inq_Master_Table_Exists_Err
    
    gSec_Inq_Master_Table_Exists = False
    
    If gSec_Inq_Master_Table(xKey) Then
        gSec_Inq_Master_Table_Exists = True
    End If

Exit Function

gSec_Inq_Master_Table_Exists_Err:
    If Err.Number = 5 Then gSec_Inq_Master_Table_Exists = False
End Function

Public Function PensionCalculation4Samuel()
    Dim cnTemp As New ADODB.Connection
    Dim dbTemp As String
    Dim qryObj As QueryDef
    Dim SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim rsStaging As New ADODB.Recordset
    Dim I, l1, L2, LR, recCount, xEmpnbr
    Dim DgDef, xMsg, Response%
    Dim K, xNum
    Dim xIncidentNo, xYear
    Dim xStr1, xStr2
    Dim empNo As Long
    Dim dblEntitle#, dblPrevEntitle#, strDivision$
    Dim strJob$, dblServiceYears#, dblServiceYearO
    Dim spt As Variant, varStartDate As Variant, lngRecs&
    Dim dblDHours#, intWhereFit&, intWhereFiO&, X%, Y%, z%, dblNewEntitle#
    Dim dblFTEHours#
    Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
    Dim Msg$, Title$
    Dim pct%
    Dim prec%, xAsOf
    Dim PenpcN, PenpcO, VED_DIV, VED_PT, SQLQW1, PenpcPre
    Dim if_Pension As Boolean
    Dim xComments
    Dim dblEntitleDays
    Dim xChange As Boolean
    Dim xChgToPay As Boolean
    Dim xAuLdate
    
    On Error GoTo Catch_Err
    
    PensionCalculation4Samuel = False
    'open Pension Master recordset for "Use Effective Service Date" only
    SQLQ = "SELECT DISTINCT PE_DIV,PE_DEPT,PE_ORG,PE_LOC,PE_SECTION,PE_EMP,PE_PT,PE_GRPCD, PE_MANUAL,PE_EDATE "
    SQLQ = SQLQ & ",PE_SALDIST " 'Ticket #22084 - Franks 05/25/2012
    SQLQ = SQLQ & ",PE_USESERVICE "
    SQLQ = SQLQ & "FROM HRPENENT "
    SQLQ = SQLQ & " WHERE NOT (PE_USESERVICE = 0) "
    
    If rsEntMain.State <> 0 Then rsEntMain.Close
    rsEntMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    I = 0
    recCount = 0
    If Not rsEntMain.EOF Then
        recCount = rsEntMain.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    
    Do While Not rsEntMain.EOF
        'spShow.FloodPercent = (I / recCount) * 100
        I = I + 1
        DoEvents
        
        'populate 1 to 24 service rules
        Call PopulatePenEntRules
        
        fglbPosGrp = ""
        Call getWSQLQ(rsEntMain)
        
        SQLQ = "SELECT ED_EMPNBR,ED_PENPCT,"
        SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
        SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "
        SQLQ = SQLQ & " ,ED_PENPCTFIXED,ED_OMERS "
        SQLQ = SQLQ & " FROM HREMP WHERE " & fglbESQLQ
        If Len(fglbPosGrp) > 0 Then
            SQLQ = SQLQ & " AND ED_EMPNBR IN "
            SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
            SQLQ = SQLQ & " WHERE JB_GRPCD = '" & fglbPosGrp & "') "
        End If
        'If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/20/2011
            SQLQ = SQLQ & "AND (ED_PENPCTFIXED IS NULL OR ED_PENPCTFIXED = 0) "
        'End If
        'SQLQ = SQLQ & " AND ED_EMPNBR=8" 'FOR TESTING
        If snapEntitle.State <> 0 Then snapEntitle.Close
        snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic
        
        If Not snapEntitle.BOF Then 'found
            'get service months

            For X% = 0 To 24
                If Not IsNumeric(xmedLTServ(X%)) Then
                    xmedLTServ(X%) = 0
                End If
                If Not IsNumeric(xmedGTServ(X%)) Then
                    xmedGTServ(X%) = 0
                Else
                   If Val(xmedGTServ(X%)) = Int(xmedGTServ(X%)) Then xmedGTServ(X%) = xmedGTServ(X%) + 0.99
                End If
                If xmedLTServ(X%) > 0 And xmedGTServ(X%) = 0 Then xmedGTServ(X%) = 9999999
            Next
            
            prec% = 0
            lngRecs& = snapEntitle.RecordCount
            'Employee Pension update - begin
            While Not snapEntitle.EOF
                prec% = prec% + 1
                pct% = Int(100 * (prec% / lngRecs&))
                MDIMain.panHelp(0).FloodPercent = pct%
                'spShow.FloodPercent = pct%
                
                if_Pension = False
            
                empNo& = snapEntitle("ED_EMPNBR")
                
              
                spt = snapEntitle("ED_PT")
                
                'If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #19938 Franks 05/26/2011
                    If Not IsNull(snapEntitle("ED_PENPCTFIXED")) Then
                        If snapEntitle("ED_PENPCTFIXED") Then
                            'If the Pension % on the Banking Information is "Fixed". Don't update it.
                            GoTo lblNextRec
                        End If
                    End If
                'End If
                
                If IsNull(snapEntitle("ED_OMERS")) Then GoTo lblNextRec
            
                varStartDate = snapEntitle("ED_OMERS")
                
               
                ''xAsOf = Date
                ''Ticket #22226 Franks 07/19/2012 - use the end date of this month to calculation the service months
                'xAsOf = CVDate(MonthName(month(Date) + 1) & " 1," & Year(Date))
                'xAsOf = DateAdd("d", -1, xAsOf)
                'Ticket #22942 Franks 12/07/2012 - the following lines caused error, no month 13
                xAsOf = CVDate(MonthName(month(Date)) & " 1," & Year(Date))
                xAsOf = DateAdd("M", 1, xAsOf) '1st day of next month
                xAsOf = DateAdd("d", -1, xAsOf)
                
                'dblServiceYears# = DateDiff("d", CVDate(varStartDate), CVDate(xAsOf)) / 365
                dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
                intWhereFit& = -1
                
                dblServiceYearO = dblServiceYears# - 1  'service until last month
                'Ticket #22509 Franks 'there was a round issue if just check from last month
                dblServiceYearO = dblServiceYears# - 2  'service until last 2 month
                
                If dblServiceYearO < 0 Then dblServiceYearO = 0
                intWhereFiO& = -1
                
                For X% = 0 To 24
                    If xmedGTServ(X%) > 0 Then
                        If dblServiceYears# >= CDbl(xmedLTServ(X%)) And dblServiceYears# <= CDbl(xmedGTServ(X%)) Then
                            intWhereFit& = X%
                            If Len(xmedPension(X%)) > 0 Then if_Pension = True
                            Exit For
                        End If
                    End If
                Next X%
                
                If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
                
                For X% = 0 To 24
                    If xmedGTServ(X%) > 0 Then
                        If dblServiceYearO >= CDbl(xmedLTServ(X%)) And dblServiceYearO <= CDbl(xmedGTServ(X%)) Then
                            intWhereFiO& = X%
                            Exit For
                        End If
                    End If
                Next X%
                
                xChange = False
                xChgToPay = False
                If if_Pension Then
                    PenpcN = xmedPension(intWhereFit&)
                    PenpcO = snapEntitle("ED_PENPCT")
                    VED_DIV = snapEntitle("ED_DIV")
                    VED_PT = snapEntitle("ED_PT")
                    'If glbFrench Then
                    '    If IsNumeric(xmedPension(intWhereFit&)) Then snapEntitle("ED_PENPCT") = Replace(xmedPension(intWhereFit&), ",", ".")
                    'Else
                        If IsNumeric(xmedPension(intWhereFit&)) Then
                            If month(varStartDate) = month(Date) Then 'Ticket #22509 Franks only update for this month
                                If IsNull(snapEntitle("ED_PENPCT")) Then
                                    xChange = True
                                Else
                                    If Not (snapEntitle("ED_PENPCT") = xmedPension(intWhereFit&)) Then
                                        xChange = True
                                    End If
                                End If
                            End If
                            
                            'Ticket #22509 Franks 09/27/2012
                            If month(varStartDate) = month(Date) Then
                                'Debug.Print "Emp#: " & snapEntitle("ED_EMPNBR") & " current pen %: " & xmedPension(intWhereFit&) & " service: " & Round(dblServiceYears#, 2)
                                 If intWhereFiO& > -1 Then
                                    PenpcPre = xmedPension(intWhereFiO&)
                                    If Not PenpcN = PenpcPre Then
                                        xChgToPay = True
                                    End If
                                 End If
                            End If
                            
                            If xChange Then
                                snapEntitle("ED_PENPCT") = xmedPension(intWhereFit&)
                            End If
                        End If
                    'End If
                    
                End If
                
                If xChange Then
                    snapEntitle.Update
                End If
                
                '''If if_Pension And xChange Then
                '''Ticket #22226 Franks 07/19/2012
                '''send the amount to hraudit even it has been the same amount on the banking screen, this is for correcting the amount in Payroll.
                ''If if_Pension Then
                ''Ticket #22509 Franks 09/06/2012 Muhammad need this for change only
                'If if_Pension And xChange Then
                'Ticket #22509 Franks 09/27/2012 - ******************
                'logic: if the Pension % in this month(PenpcN) <> the Pension % in last month(PenpcPre) from the calculation
                '       then send the PenpcN to Audit anyway, no matter if the Pension % has been PenpcN
                If if_Pension And xChgToPay Then
                '****************************************************************************
                    If month(varStartDate) = month(Date) Then
                    'Ticket #22226 Franks 06/28/2012 only write data to audit table for current month
                        ''PensionCalculation4Samuel = True 'updated
                        'Ticket #22226 Franks 06/27/2012
                        'Muhammad asked to change: alway pass the 1st day of following month
                        'If Day(Date) < 15 Then
                        '    xAuLdate = CVDate(MonthName(month(Date)) & " 1," & Year(Date))
                        'Else
                            'xAuLdate = CVDate(MonthName(month(Date) + 1) & " 1," & Year(Date))
                            'Ticket #22942 Franks 12/07/2012 - the following lines caused error, no month 13
                            xAuLdate = CVDate(MonthName(month(Date)) & " 1," & Year(Date))
                            xAuLdate = DateAdd("M", 1, xAuLdate) '1st day of next month
                        'End If
                        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_PENPCT,AU_OLDPEN, "
                        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "
                        
                        SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo& & "," & Val(Format(PenpcN)) & "," & Val(Format(PenpcO))
                        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
                        'SQLQW1 = SQLQW1 & Date_SQL(Date) & ", '"
                        SQLQW1 = SQLQW1 & Date_SQL(xAuLdate) & ", '"
                        SQLQW1 = SQLQW1 & Time$ & "', "
                        SQLQW1 = SQLQW1 & "'N', "
                        SQLQW1 = SQLQW1 & "'PenTask'"
                        SQLQW1 = SQLQW1 & ")"
                        gdbAdoIhr001X.Execute SQLQW1
                    End If
                End If
            
lblNextRec:
                snapEntitle.MoveNext
                DoEvents
            Wend
                        
            'Employee Pension Update - end
        
        End If
        snapEntitle.Close
        
        rsEntMain.MoveNext
    Loop
    rsEntMain.Close
    PensionCalculation4Samuel = True
    
    MDIMain.panHelp(0).FloodType = 0
    
    Exit Function
    
Catch_Err:

    PensionCalculation4Samuel = False
    
End Function

Sub PopulatePenEntRules()
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim X
For X = 0 To 24
    xmedLTServ(X) = ""
    xmedGTServ(X) = ""
    xmedPension(X) = ""
Next

If Not rsEntMain.EOF Then
    SQLQ = "SELECT * FROM HRPENENT WHERE (1=1) "
    If IsNull(rsEntMain("PE_DIV")) Then
        SQLQ = SQLQ & " AND PE_DIV IS NULL"
    Else
        If Len(rsEntMain("PE_DIV")) > 0 Then SQLQ = SQLQ & " AND PE_DIV = '" & rsEntMain("PE_DIV") & "'"
    End If
    If IsNull(rsEntMain("PE_DEPT")) Then
        SQLQ = SQLQ & " AND PE_DEPT IS NULL"
    Else
        If Len(rsEntMain("PE_DEPT")) > 0 Then SQLQ = SQLQ & " AND PE_DEPT = '" & rsEntMain("PE_DEPT") & "'"
    End If
    If IsNull(rsEntMain("PE_ORG")) Then
        SQLQ = SQLQ & " AND PE_ORG IS NULL"
    Else
        If Len(rsEntMain("PE_ORG")) > 0 Then SQLQ = SQLQ & " AND PE_ORG = '" & rsEntMain("PE_ORG") & "'"
    End If
    If IsNull(rsEntMain("PE_LOC")) Then
        SQLQ = SQLQ & " AND PE_LOC IS NULL"
    Else
        If Len(rsEntMain("PE_LOC")) > 0 Then SQLQ = SQLQ & " AND PE_LOC = '" & rsEntMain("PE_LOC") & "'"
    End If
    If IsNull(rsEntMain("PE_SECTION")) Then
        SQLQ = SQLQ & " AND PE_SECTION IS NULL"
    Else
        If Len(rsEntMain("PE_SECTION")) > 0 Then SQLQ = SQLQ & " AND PE_SECTION = '" & rsEntMain("PE_SECTION") & "'"
    End If
    If IsNull(rsEntMain("PE_EMP")) Then
        SQLQ = SQLQ & " AND PE_EMP IS NULL"
    Else
        If Len(rsEntMain("PE_EMP")) > 0 Then SQLQ = SQLQ & " AND PE_EMP = '" & rsEntMain("PE_EMP") & "'"
    End If
    If IsNull(rsEntMain("PE_PT")) Then
        SQLQ = SQLQ & " AND PE_PT IS NULL"
    Else
        If Len(rsEntMain("PE_PT")) > 0 Then SQLQ = SQLQ & " AND PE_PT = '" & rsEntMain("PE_PT") & "' "
    End If
    'Ticket #22084 - Franks 05/25/2012
    If IsNull(rsEntMain("PE_SALDIST")) Then
        SQLQ = SQLQ & " AND PE_SALDIST IS NULL"
    Else
        If Len(rsEntMain("PE_SALDIST")) > 0 Then SQLQ = SQLQ & " AND PE_SALDIST = '" & rsEntMain("PE_SALDIST") & "' "
    End If
    If IsNull(rsEntMain("PE_GRPCD")) Then
        SQLQ = SQLQ & " AND PE_GRPCD IS NULL"
    Else
        If Len(rsEntMain("PE_GRPCD")) > 0 Then SQLQ = SQLQ & " AND PE_GRPCD = '" & rsEntMain("PE_GRPCD") & "'"
    End If

    
    SQLQ = SQLQ & " Order By PE_DIV,PE_DEPT,PE_ORG, PE_EDATE,PE_EMP,PE_PT,PE_LOC,PE_SECTION,PE_SALDIST,PE_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    
    Do While Not rsVE.EOF
        xOrder = rsVE("PE_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 24) Then
            If Not IsNull(rsVE("PE_BMONTH")) Then xmedLTServ(nOrder) = rsVE("PE_BMONTH")
            If Not IsNull(rsVE("PE_EMONTH")) Then xmedGTServ(nOrder) = rsVE("PE_EMONTH")
            If Not IsNull(rsVE("PE_PCT")) Then xmedPension(nOrder) = rsVE("PE_PCT")
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close
End If

End Sub

Private Sub getWSQLQ(rsEnt As ADODB.Recordset)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate

'SQLQ = "SELECT DISTINCT PE_DIV,PE_DEPT,PE_ORG,PE_LOC,PE_SECTION,PE_EMP,PE_PT,PE_GRPCD, PE_MANUAL,PE_EDATE "

fglbESQLQ = " (1=1) " 'glbSeleDeptUn
If Not IsNull(rsEnt("PE_DEPT")) Then
    If Len(rsEnt("PE_DEPT")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & rsEnt("PE_DEPT") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_DIV")) Then
    If Len(rsEnt("PE_DIV")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & rsEnt("PE_DIV") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_ORG")) Then
    If Len(rsEnt("PE_ORG")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & rsEnt("PE_ORG") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_EMP")) Then
    If Len(rsEnt("PE_EMP")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & rsEnt("PE_EMP") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_SECTION")) Then
    If Len(rsEnt("PE_SECTION")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & rsEnt("PE_SECTION") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_LOC")) Then
    If Len(rsEnt("PE_LOC")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & rsEnt("PE_LOC") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_PT")) Then
    If Len(rsEnt("PE_PT")) > 0 Then
        fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & rsEnt("PE_PT") & "' "
    End If
End If
If Not IsNull(rsEnt("PE_GRPCD")) Then
    If Len(rsEnt("PE_GRPCD")) > 0 Then
        fglbPosGrp = rsEnt("PE_GRPCD")
    End If
End If

'Ticket #22084 - Franks 05/25/2012
If glbCompSerial = "S/N - 2382W" Then
    If Not IsNull(rsEnt("PE_SALDIST")) Then
        If Len(rsEnt("PE_SALDIST")) > 0 Then
            fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & rsEnt("PE_SALDIST") & "' "
        End If
    End If
End If

End Sub

