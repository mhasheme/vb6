VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOutLookAddress 
   Caption         =   "Select Recipients"
   ClientHeight    =   4188
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10188
   LinkTopic       =   "Form1"
   ScaleHeight     =   4188
   ScaleWidth      =   10188
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBCCRemove 
      Caption         =   "< -- Bcc"
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCCRemove 
      Caption         =   "< -- Cc"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdToRemove 
      Caption         =   "< -- To "
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstBCC 
      Height          =   816
      Left            =   6240
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ListBox lstCC 
      Height          =   816
      Left            =   6240
      TabIndex        =   8
      Top             =   1680
      Width           =   3735
   End
   Begin VB.ListBox lstTo 
      Height          =   816
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   3735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdBcc 
      Caption         =   "Bcc -- >"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCC 
      Caption         =   "Cc --  >"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdTo 
      Caption         =   "To -- >"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      _ExtentX        =   7853
      _ExtentY        =   5313
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblMesRec 
      Caption         =   "Message Recipients"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmOutLookAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim olApp As Outlook.Application
Dim olNS As Outlook.NameSpace
Dim olAL As Outlook.AddressList
Dim olAE As Outlook.AddressEntry
Dim olMail As Outlook.MailItem

Private Sub cmdBcc_Click()
    Dim iItem
    For iItem = 1 To lvw.ListItems.count
        If lvw.ListItems(iItem).selected Then lstBCC.AddItem lvw.ListItems(iItem).ListSubItems.Item(1).Text
        
    Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCC_Click()
    Dim iItem
    For iItem = 1 To lvw.ListItems.count
        If lvw.ListItems(iItem).selected Then lstCC.AddItem lvw.ListItems(iItem).ListSubItems.Item(1).Text
        
    Next
'    lstCC.AddItem lvw.SelectedItem.ListSubItems.Item(1).Text
End Sub

Private Sub cmdOK_Click()
    Dim iItem
    If lstTo.ListCount > 0 Then
        For iItem = 0 To lstTo.ListCount - 1
            If Len(frmSendEmail.txtTo.Text) > 0 Then
                frmSendEmail.txtTo.Text = frmSendEmail.txtTo.Text & ";" & lstTo.List(iItem)
            Else
                frmSendEmail.txtTo.Text = lstTo.List(iItem)
            End If
        Next
    End If
    If lstCC.ListCount > 0 Then
        For iItem = 0 To lstCC.ListCount - 1
            If Len(frmSendEmail.txtCC.Text) > 0 Then
                frmSendEmail.txtCC.Text = frmSendEmail.txtCC.Text & ";" & lstCC.List(iItem)
            Else
                frmSendEmail.txtCC.Text = lstCC.List(iItem)
            End If
        Next
    End If
'    If lstBCC.ListCount > 0 Then
'        For iItem = 1 To lstBCC.ListCount
'            If Len(frmSendEmail.txtBCC.Text) > 0 Then
'                frmSendEmail.txtBCC.Text = frmSendEmail.txtTo.Text & ";" & lstBCC.List(iItem)
'            Else
'                frmSendEmail.txtBCC.Text = lstBCC.List(iItem)
'            End If
'        Next
'    End If
    Unload Me
End Sub

Private Sub cmdTo_Click()
    
    Dim iItem
    For iItem = 1 To lvw.ListItems.count
        If lvw.ListItems(iItem).selected Then lstTo.AddItem lvw.ListItems(iItem).ListSubItems.Item(1).Text
        
    Next
'    lstTo.AddItem lvw.SelectedItem.ListSubItems.Item(1).Text
End Sub

Private Sub cmdToRemove_Click()
    lstTo.RemoveItem lstTo.ListIndex
End Sub

Private Sub cmdCCRemove_Click()
    lstCC.RemoveItem lstCC.ListIndex
End Sub

Private Sub Form_Load()
On Error Resume Next
  Set olApp = New Outlook.Application
  Set olNS = olApp.GetNamespace("MAPI")
  For Each olAL In olNS.AddressLists
'    If IsObject(olAL.AddressEntries) Then
    For Each olAE In olAL.AddressEntries
     If IsValidEmail(olAE.Address) = True Then
        lvw.ListItems.Add , , olAE.name
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = olAE.Address
        'lvw.ListItems(lvw.ListItems.count).SubItems(2) = olAE.ID
        lvw.ListItems(lvw.ListItems.count).Tag = olAE.ID
      End If
    Next
'    Else
    If Err.Number = 287 Then GoTo NoAddress
'    End If
  Next
NoAddress:
    'Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  olApp.Quit
  Set olApp = Nothing
End Sub

Private Sub lstTo_Click()
    cmdTo.Visible = False
    cmdToRemove.Visible = True
End Sub

Private Sub lstCC_Click()
    cmdCC.Visible = False
    cmdCCRemove.Visible = True
End Sub

Private Sub lvw_Click()
    cmdTo.Visible = True
    cmdToRemove.Visible = False
    cmdCC.Visible = True
    cmdCCRemove.Visible = False
    
End Sub

Private Sub lvw_DblClick()
If Len(frmSendEmail.txtTo.Text) > 0 Then
    frmSendEmail.txtTo.Text = frmSendEmail.txtTo.Text & ";" & lvw.SelectedItem.ListSubItems.Item(1).Text
Else
    frmSendEmail.txtTo.Text = lvw.SelectedItem.ListSubItems.Item(1).Text
End If
'frmSendEmail.txtTo.Text = lvw.SelectedItem.Text 'lvw.SelectedItem.ListSubItems.Item(1).Text
Unload Me
End Sub


Public Function IsValidEmail(email As String) As Boolean
    Dim myAt As Integer
    Dim myAtLastPos As Integer
    Dim myDot As Integer
    Dim myDotDot As Integer
    Dim myDotAt As Integer
    Dim myAtDot As Integer
    Dim mySpace As Integer
    IsValidEmail = True
    mySpace = InStr(1, email, " ", vbTextCompare)
    myAtLastPos = InStrRev(email, "@", , vbTextCompare)
    myAt = InStr(1, email, "@", vbTextCompare)
    myAtDot = InStr(1, email, "@.", vbTextCompare)
    myDotAt = InStr(1, email, ".@", vbTextCompare)
    myDot = InStr(myAt + 2, email, ".", vbTextCompare)
    myDotDot = InStr(myAt + 2, email, "..", vbTextCompare)
    If myAtDot > 0 Or myDotAt > 0 Or myAtLastPos <> myAt Or mySpace > 0 Or myAt = 0 Or myDot = 0 Or myDotDot > 0 Or Right(email, 1) = "." Then IsValidEmail = False
End Function

